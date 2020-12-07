using System;

namespace App
{
    public class RepostPage
    {
        public string Title { get; set; }
        public string Description { get; set; }
        public string BannerImageUrl { get; set; }
        public string OriginalSourceUrl { get; set; }
        public int? Id { get; set; } 
        public string FileName { get; set; }
    }

    public class RepostPageCreationInfo : RepostPage
    {
        public bool IsBannerImageUrlExternal { get; set; }
        public bool ShouldSaveAsDraft { get; set; }
    }
    public class RepostPageItem : RepostPage
    {
        public RepostPageItem(RepostPage repostPage)
        {
            this.Id = repostPage.Id ?? -1;
            this.FileName = repostPage.FileName ?? "";
            this.Title = repostPage.Title;
            this.Description = repostPage.Description;
            this.BannerImageUrl = repostPage.BannerImageUrl;
            this.OriginalSourceUrl = repostPage.OriginalSourceUrl;
        }
        public string SourcePostSourceType { get; set; }
        public string SourceProvider { get; set; }
        public string SourceCategories { get; set; }
        public DateTime SourceModifiedDate { get; set; }
        public DateTime SourcePublishDate { get; set; }
    }

    public class RepostPageCollection
    {
        public RepostPage[] value { get; set; }
    }
}
