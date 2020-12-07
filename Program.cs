using Microsoft.Extensions.Configuration;
using Microsoft.SharePoint.Client;
using System;
using System.Linq;
using System.Threading.Tasks;
using System.Security.Cryptography.X509Certificates;
using System.Text.Json;
using PnP.Framework;
using PnP.Framework.Pages;
using System.IO;
using System.Web;
using Microsoft.SharePoint.Client.Taxonomy;
using System.Collections.Generic;

namespace App
{
  class Program
  {
    private static AppConfig CONFIG = null;
    private static Dictionary<string, Field> MEM_CACHE_FIELDS = new Dictionary<string, Field>();
    private static Dictionary<string, Term> MEM_CACHE_TERMS = new Dictionary<string, Term>();
    private static Dictionary<string, User> MEM_CACHE_GROUPS = new Dictionary<string, User>();

    public static async Task Main(string[] args)
    {
      try
      {
        LogInfo("SharePoints External News Links Processor Started");

        Program.CONFIG = LoadAppSettings();

        if (Program.CONFIG == null)
          return;

        var certificate = GetCertificateFromStore(Program.CONFIG.CertificateThumbprint);

        var newsPosts = await GetNewsPosts();
        LogInfo($"Found {newsPosts.Length} news posts");
        if (newsPosts.Length == 0) return;

        using (var authManager = new AuthenticationManager(
            clientId: Program.CONFIG.ClientId,
            certificate: certificate,
            tenantId: Program.CONFIG.TenantId
        ))
        {
          var clientContext = await authManager.GetContextAsync(Program.CONFIG.SiteUrl);
          var context = PnPClientContext.ConvertFrom(clientContext);

          var web = context.Web;
          context.Load(web, w => w.Title, w => w.Url);
          await context.ExecuteQueryAsync();

          LogInfo($"Connected to '{web.Title}' at {web.Url}");

          foreach (var newsPost in newsPosts)
          {
            try
            {
              LogInfo($"Processing news post: {newsPost?.Title}");
              AddOrUpdateRepostPage(context, newsPost);
            }
            catch (Exception ex)
            {
              LogError($"Error processing news post {newsPost?.Title}", ex);
            }
            finally
            {
              LogInfo($"Finished processing news post");
            }
          }
        }
      }
      catch (Exception ex)
      {
        LogError("Unhandled application exception", ex);
      }
      
      LogInfo("SharePoints External News Links Processor Finished");
    }

    private static void AddOrUpdateRepostPage(PnPClientContext context, NewsPost post)
    {
      string pageName = $"{post.UrlSlug}.aspx";

      var repostPage = EnsureRepostPage(context, pageName);

      string postUrl = post.PermaLink ?? String.Format(Program.CONFIG.CustomNewsLinkStringTemplate, post.PostId);
      if (postUrl.Length > 255)
      {
        LogWarning("  Repost URL exceeds the maximum allowed length of 255 characters. URL will be truncated to 255 characters.");
        postUrl = postUrl.Substring(0, 255);
      }
      string imageUrl = post.Images.Box960.Url;

      bool hasChanges = false;
      ListItem repostItem = repostPage.PageListItem;

      // Audience Targeting
      var audienceGroupNames = MapCategoriesToGroupNames(post.Categories.Select(c => c.Name).Distinct().ToArray());
      var resolvedGroups = ResolveGroups(context, audienceGroupNames);
      var groupFieldValues = resolvedGroups.Select(rg => new FieldUserValue() { LookupId = rg.Id }).ToList();
      if (SetFieldValue(ref repostItem, "_ModernAudienceTargetUserField", groupFieldValues)) hasChanges = true;      

      // Set Custom Repost Content Type
      if (SetFieldValue(ref repostItem, ClientSidePage.ContentTypeId, Program.CONFIG.CustomRepostContentTypeId)) hasChanges = true;

      // Set Core Fields 
      if (SetFieldValue(ref repostItem, ClientSidePage.Title, post.Title)) hasChanges = true;
      if (SetFieldValue(ref repostItem, ClientSidePage.DescriptionField, post.Description)) hasChanges = true;
      if (SetFieldValue(ref repostItem, ClientSidePage.FirstPublishedDate, DateTime.UtcNow, overwrite: false)) hasChanges = true;
      if (SetFieldValue(ref repostItem, ClientSidePage.PromotedStateField, (int)PromotedState.Promoted)) hasChanges = true;
      if (SetFieldValue(ref repostItem, ClientSidePage._OriginalSourceUrl, postUrl)) hasChanges = true;
      if (SetFieldValue(ref repostItem, ClientSidePage._OriginalSourceSiteId, Guid.Empty)) hasChanges = true;
      if (SetFieldValue(ref repostItem, ClientSidePage._OriginalSourceWebId, Guid.Empty)) hasChanges = true;
      if (SetFieldValue(ref repostItem, ClientSidePage._OriginalSourceListId, Guid.Empty)) hasChanges = true;
      if (SetFieldValue(ref repostItem, ClientSidePage._OriginalSourceItemId, Guid.Empty)) hasChanges = true;
      if (SetFieldValue(ref repostItem, ClientSidePage.BannerImageUrl, new FieldUrlValue() { Description = imageUrl, Url = imageUrl })) hasChanges = true;

      // Set Core Repost Content
      string layoutsWebpartContent = GetRepostLayoutWebpartsContent(postUrl, imageUrl, post);
      if (SetFieldValue(ref repostItem, ClientSidePage.PageLayoutContentField, layoutsWebpartContent)) hasChanges = true;
      if (hasChanges) 
      {
        LogInfo("  Updating news link core fields");
        repostItem.Update();
        context.ExecuteQueryRetry();
      }

      // Set Extra Metadata
      var terms = GetManagedMetadataTerms(context, post.Categories.Select(c => c.Name).Distinct().ToArray());
      if (SetFieldValue(ref repostItem, "SourceCategories", terms)) hasChanges = true;
      if (SetFieldValue(ref repostItem, "SourcePostSourceType", post.PostSourceType)) hasChanges = true;
      if (SetFieldValue(ref repostItem, "SourceProvider", post.Provider)) hasChanges = true;
      if (SetFieldValue(ref repostItem, "SourceModifiedDate", post.ModifiedDate)) hasChanges = true;
      if (SetFieldValue(ref repostItem, "SourcePublishDate", post.PublishDate)) hasChanges = true;

      if (hasChanges)
      {
        LogInfo("  Updating news link metadata");
        repostItem.Update();
        context.ExecuteQueryRetry();
        repostPage.Publish();
      }
      else
      {
        LogInfo("  News link is up-to-date; no changes made");
      }
    }

    private static ClientSidePage EnsureRepostPage(PnPClientContext context, string pageName)
    {
      ClientSidePage repostPage = null;
      try
      {
        // Check if the repost already exists
        repostPage = ClientSidePage.Load(context, pageName);
        LogInfo($"  Found existing news link: {pageName}");
      }
      catch (ArgumentException ex)
      {
        if (ex.Message.Contains("does not exist"))
        {
          // If it doesn't, create it
          repostPage = context.Web.AddClientSidePage(pageName);
          repostPage.LayoutType = ClientSidePageLayoutType.RepostPage;
          repostPage.Save(pageName);
          repostPage = ClientSidePage.Load(context, pageName);
          LogInfo($"  Created news link: {pageName}");
        }
        else throw;
      }
      catch (Exception ex)
      {
        LogError("Unable to retrieve or add repost page", ex);
      }
      return repostPage;
    }

    private static string[] MapCategoriesToGroupNames(string[] categories)
    {
      List<string> audiences = new List<string>();

      foreach (var category in categories)
      {
        switch (category)
        {
          case "Engineering":
            audiences.Add("news-engineering"); break;

          case "Our News": 
          case "Firm in the News": 
          case "Firm & Industry News": 
            audiences.Add("news-company"); break;

          case "Health & Wellness": 
            audiences.Add("news-health"); break;

          case "Careers": 
            audiences.Add("news-careers"); break;
        }
      }

      return audiences.Distinct().ToArray();
    }

    private static string GetRepostLayoutWebpartsContent(string postUrl, string imageUrl, NewsPost post)
    {
      string outerTemplate = @"<div><div data-sp-canvascontrol="""" data-sp-canvasdataversion="""" data-sp-controldata=""{0}""></div></div>";
      var controlData = new
      {
        id = "c1b5736d-84dd-4fdb-a7be-e7e9037bd3c3",
        instanceId = "c1b5736d-84dd-4fdb-a7be-e7e9037bd3c3",
        serverProcessedContent = new
        {
          htmlStrings = new { },
          searchablePlainTexts = new { },
          imageSources = new { },
          links = new { },
        },
        dataVersion = "1.0",
        properties = new
        {
          description = post.Description,
          thumbnailImageUrl = imageUrl,
          title = post.Title,
          url = postUrl
        }
      };
      string controlDataJson = JsonSerializer.Serialize(controlData);
      string encodedControlData = HttpUtility.HtmlEncode(controlDataJson);
      encodedControlData = encodedControlData.Replace("{", "&#123;");
      encodedControlData = encodedControlData.Replace("}", "&#125;");
      encodedControlData = encodedControlData.Replace(":", "&#58;");

      return string.Format(outerTemplate, encodedControlData);
    }

    private static bool SetFieldValue(ref ListItem item, string fieldName, DateTime fieldValue, bool overwrite = true)
    {
      var hasChanged = false;
      var itemValue = item.GetFieldValueAs<DateTime>(fieldName);
      if ((itemValue == DateTime.MinValue && fieldValue != null && fieldValue != DateTime.MinValue)
          || (overwrite && !String.Format("{0:u}", itemValue.ToUniversalTime()).Equals(String.Format("{0:u}", fieldValue.ToUniversalTime()))))
      {
        item[fieldName] = fieldValue;
        hasChanged = true;
      }
      return hasChanged;
    }

    private static bool SetFieldValue(ref ListItem item, string fieldName, List<Term> fieldValue, bool overwrite = true)
    {
      var hasChanged = false;
      var itemValue = item[fieldName] as TaxonomyFieldValueCollection;
      var itemTermLabels = (from term in itemValue orderby term.Label select term.Label).ToArray();
      var fieldTermLabels = (from term in fieldValue orderby term.Name select term.Name).ToArray();

      if ((itemValue == null && fieldValue != null) || (overwrite && !itemTermLabels.SequenceEqual(fieldTermLabels)))
      {
        var field = GetFieldFromInternalName(item, fieldName);
        string termValuesString = String.Empty;
        foreach(var term in fieldValue)
        {
          termValuesString += "-1;#" + term.Name + "|" + term.Id.ToString("D") + ";#";
        }
        termValuesString = termValuesString.Substring(0, termValuesString.Length - 2);
        var termsFieldValue = new TaxonomyFieldValueCollection(item.Context, termValuesString, field);
        item[fieldName] = termsFieldValue;
        hasChanged = true;
      }
      return hasChanged;
    }

    private static bool SetFieldValue(ref ListItem item, string fieldName, List<FieldUserValue> fieldValue, bool overwrite = true)
    {
      var hasChanged = false;
      var itemValue = item[fieldName] as FieldUserValue[];
      var itemGroupIds = itemValue != null ? (from grp in itemValue orderby grp.LookupId select grp.LookupId).ToArray() : null;
      var fieldGroupIds = fieldValue != null && fieldValue.Count > 0 ? (from grp in fieldValue orderby grp.LookupId select grp.LookupId).ToArray() : null;

      if ((itemValue == null && fieldValue != null && fieldValue.Count > 0) 
          || (overwrite && itemGroupIds != null && fieldGroupIds != null && !itemGroupIds.SequenceEqual(fieldGroupIds)))
      {
        item[fieldName] = fieldValue;
        hasChanged = true;
      }
      return hasChanged;
    }

    private static bool SetFieldValue(ref ListItem item, string fieldName, FieldUrlValue fieldValue, bool overwrite = true)
    {
      var hasChanged = false;
      var itemValue = item[fieldName] as FieldUrlValue;
      if ((itemValue == null && fieldValue != null)
          || (overwrite && (!itemValue.Url.Equals(fieldValue.Url) || !itemValue.Description.Equals(fieldValue.Description))))
      {
        item[fieldName] = fieldValue;
        hasChanged = true;
      }
      return hasChanged;
    }

    private static bool SetFieldValue<T>(ref ListItem item, string fieldName, T fieldValue, bool overwrite = true)
    {
      var hasChanged = false;
      var itemValue = item[fieldName];
      if ((itemValue == null && fieldValue != null) || (overwrite && !itemValue.ToString().Equals(fieldValue.ToString())))
      {
        item[fieldName] = fieldValue;
        hasChanged = true;
      }
      return hasChanged;
    }

    private static User[] ResolveGroups(PnPClientContext context, string[] groupNames)
    {
      List<User> groups = new List<User>();
      
      foreach (var groupName in groupNames)
      {
        var groupCache = MEM_CACHE_GROUPS.FirstOrDefault(g => g.Key == groupName);
        if (groupCache.Value == null)
        {
          try 
          {
            var group = context.Web.EnsureUser(groupName);
            context.Load(group);
            context.ExecuteQueryRetry();
            groups.Add(group);
            MEM_CACHE_GROUPS.Add(groupName, group);
          }
          catch (Exception ex)
          {
            LogError($"Unable to resolve group: {groupName}", ex);
          }
        }
        else
        {
          groups.Add(groupCache.Value);
        }
      }
      return groups.ToArray();
    }

    private static List<Term> GetManagedMetadataTerms(PnPClientContext context, string[] termLabels)
    {
      List<Term> resolvedTerms = new List<Term>();
      foreach (var label in termLabels)
      {
        var termCache = MEM_CACHE_TERMS.FirstOrDefault(t => t.Key == label);
        if (termCache.Value == null)
        {
          try
          {
            Term term = context.Site.GetTermByName(Guid.Parse(Program.CONFIG.CustomCategoriesTermSetId), label);
            if (term != null)
            {
              resolvedTerms.Add(term);
              MEM_CACHE_TERMS.Add(label, term);
            }
          }
          catch (Exception ex)
          {
            LogError($"Error fetching managed metadata term: {label}", ex);
          }
        }
        else
        {
          resolvedTerms.Add(termCache.Value);
        }
      }
      return resolvedTerms;
    }

    private static Field GetFieldFromInternalName(ListItem item, string fieldName)
    {
      var fieldIdCache = MEM_CACHE_FIELDS.FirstOrDefault(f => f.Key == fieldName);
      if (fieldIdCache.Value == null)
      {
        var field = item.ParentList.Fields.GetByInternalNameOrTitle(fieldName);
        item.Context.Load(field, f => f.Id);
        item.Context.ExecuteQueryRetry();
        MEM_CACHE_FIELDS.Add(fieldName, field);
        return field;
      }
      else return fieldIdCache.Value;
    }

    private static async Task<NewsPost[]> GetNewsPosts()
    {
      try
      {
        using (StreamReader file = System.IO.File.OpenText("sampledata.json"))
        {
          var sampleData = await file.ReadToEndAsync();
          var opts = new JsonSerializerOptions()
          {
            PropertyNameCaseInsensitive = true
          };
          var newsPosts = JsonSerializer.Deserialize<NewsPostCollection>(sampleData, opts);
          // return new NewsPost[] { newsPosts.posts.ToArray().Last() }; //.Skip(5).Take(5)
          return newsPosts.posts.ToArray(); //.Skip(5).Take(5)
        }
      }
      catch (Exception ex)
      {
        LogError("Unable to load sample data file", ex);
        return null;
      }
    }

    private static AppConfig LoadAppSettings()
    {
      AppConfig appConfig = new AppConfig();
      try
      {
        var userSecrets = new ConfigurationBuilder()
            .AddUserSecrets<Program>()
            .Build();

        // Core Authentication and Site Configuration
        appConfig.ClientId = userSecrets["ClientId"];
        appConfig.TenantId = userSecrets["TenantId"];
        appConfig.SiteUrl = userSecrets["SiteUrl"];
        appConfig.CertificateThumbprint = userSecrets["CertificateThumbprint"];

        // Custom Configuration
        appConfig.CustomRepostContentTypeId = userSecrets["CustomRepostContentTypeId"];
        appConfig.CustomCategoriesTermSetId = userSecrets["CustomCategoriesTermSetId"];
        appConfig.CustomNewsLinkStringTemplate = userSecrets["CustomNewsLinkStringTemplate"];
      }
      catch (Exception ex)
      {
        LogError("Unable to load app configuration", ex);
        appConfig = null;
      }

      return appConfig;
    }

    private static X509Certificate2 GetCertificateFromStore(string thumbprint, StoreName storeName = StoreName.My, StoreLocation storeLocation = StoreLocation.CurrentUser)
    {

      X509Store store = new X509Store(storeName, storeLocation);
      try
      {
        store.Open(OpenFlags.ReadOnly | OpenFlags.OpenExistingOnly);
        X509Certificate2Collection certificates = store.Certificates.Find(
            X509FindType.FindByThumbprint, thumbprint, false);
        if (certificates.Count == 1)
        {
          return certificates[0];
        }
        else
        {
          return null;
        }
      }
      finally
      {
        store.Close();
      }
    }

    private static void LogError(string message, Exception exception = null)
    {
      Console.ForegroundColor = ConsoleColor.Red;
      Console.WriteLine($"[{String.Format("{0:u}", DateTime.Now)}] {message}");
      if (exception != null)
      {
        Console.WriteLine(exception);
      }
      Console.ResetColor();
    }

    private static void LogWarning(string message)
    {
      Console.ForegroundColor = ConsoleColor.Yellow;
      Console.WriteLine($"[{String.Format("{0:u}", DateTime.Now)}] {message}");
      Console.ResetColor();
    }

    private static void LogInfo(string message)
    {
      Console.ForegroundColor = ConsoleColor.White;
      Console.WriteLine($"[{String.Format("{0:u}", DateTime.Now)}] {message}");
      Console.ResetColor();
    }
  }
}