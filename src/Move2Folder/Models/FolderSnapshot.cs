using System.Collections.Generic;

namespace Move2Folder.Models
{
    public class FolderSnapshot
    {
        public string EntryId { get; set; } = string.Empty;
        public string StoreId { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
        public string FullPath { get; set; } = string.Empty;
        public string DisplayPath { get; set; } = string.Empty;
        public bool IsStore { get; set; }
        public string FolderType { get; set; } = "Folder";
        public List<FolderSnapshot> Children { get; } = new List<FolderSnapshot>();
    }
}
