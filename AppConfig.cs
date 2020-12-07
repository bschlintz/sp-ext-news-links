using System;

namespace App
{
    public class AppConfig
    {
        // Core Authentication and Site Configuration
        public string ClientId { get; set; }
        public string TenantId { get; set; }
        public string CertificateThumbprint { get; set; }
        public string SiteUrl { get; set; }

        // Custom Configuration
        public string CustomRepostContentTypeId { get; set; }
        public string CustomNewsLinkStringTemplate { get; set; }
        public string CustomCategoriesTermSetId { get; set; }

    }
}
