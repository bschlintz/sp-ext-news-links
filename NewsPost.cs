using System;

namespace App
{
    public class NewsPost
    {
        public string PostId { get; set; }
        public string UrlSlug { get; set; }
        public string Title { get; set; }
        public string Description { get; set; }
        public string PermaLink { get; set; }
        public string PostSourceType { get; set; }
        public string Provider { get; set; }
        public DateTime ModifiedDate { get; set; }
        public DateTime PublishDate { get; set; }
        public NewsPostCategory[] Categories { get; set; }
        public NewsPostImageCollection Images { get; set; }
    }

    public class NewsPostImageCollection
    {
        public NewsPostImage Original { get; set; }
        public NewsPostImage Box1440 { get; set; }
        public NewsPostImage Box1280 { get; set; }
        public NewsPostImage Box960 { get; set; }
        public NewsPostImage Box640 { get; set; }
        public NewsPostImage Box320 { get; set; }
    }

    public class NewsPostImage
    {
        public string Size { get; set; }
        public string Url { get; set; }
        public int Width { get; set; }
        public int Height { get; set; }
        public string MimeType { get; set; }
    }

    public class NewsPostCategory
    {
        public int Id { get; set; }
        public int? ParentCategoryId { get; set; }        
        public string Name { get; set; }
    }

    public class NewsPostCollection
    {
        public NewsPost[] posts { get; set; }
    }
}