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