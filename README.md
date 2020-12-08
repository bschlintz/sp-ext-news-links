# SharePoint External News Link Importer
Import news links to an SPO Site Pages library using PnP Core.

## Configuration
1. Open your command line interface (CLI) in the directory where **sp-ext-news-links.csproj** is located.
1. Run the following command to initialize [user secrets](https://docs.microsoft.com/aspnet/core/security/app-secrets) for the project.

    ```dotnetcli
    dotnet user-secrets init
    ```

1. Run the following commands to store your app ID, app secret, and tenant ID in the user secret store.

    ```dotnetcli
    dotnet user-secrets set ClientId "YOUR_APP_ID"
    dotnet user-secrets set TenantId "YOUR_TENANT_ID"
    dotnet user-secrets set SiteUrl "YOUR_SITE_URL"
    dotnet user-secrets set CertificateThumbprint "YOUR_CERTIFICATE_THUMBPRINT"
    ```
    
1. Run the following to set use case-specific values in the user secret store.
    ```dotnetcli
    dotnet user-secrets set CustomRepostContentTypeId "YOUR_REPOST_CONTENT_TYPE_ID"
    dotnet user-secrets set CustomCategoriesTermSetId "YOUR_CATEGORIES_TERM_SET_ID"
    dotnet user-secrets set CustomNewsLinkStringTemplate "YOUR_NEWS_LINK_STRING_TEMPLATE"
    ```

## Disclaimer

Microsoft provides programming examples for illustration only, without warranty either expressed or implied, including, but not limited to, the implied warranties of merchantability and/or fitness for a particular purpose. We grant You a nonexclusive, royalty-free right to use and modify the Sample Code and to reproduce and distribute the object code form of the Sample Code, provided that You agree: (i) to not use Our name, logo, or trademarks to market Your software product in which the Sample Code is embedded; (ii) to include a valid copyright notice on Your software product in which the Sample Code is embedded; and (iii) to indemnify, hold harmless, and defend Us and Our suppliers from and against any claims or lawsuits, including attorneys' fees, that arise or result from the use or distribution of the Sample Code.