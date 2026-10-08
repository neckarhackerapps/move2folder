using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace Move2Folder.Models
{
    public class FolderItem : INotifyPropertyChanged
    {
        private bool _isExpanded;
        private bool _isSelected;
        private bool _isMatching;

        public string EntryId { get; set; } = string.Empty;
        public string StoreId { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
        public string FullPath { get; set; } = string.Empty;
        public string DisplayPath { get; set; } = string.Empty;
        public bool IsStore { get; set; }
        public string FolderType { get; set; } = "Folder"; // "Store", "Inbox", "Sent", "Archive", "Folder", etc.

        public FolderItem? Parent { get; set; }
        public ObservableCollection<FolderItem> Children { get; } = new ObservableCollection<FolderItem>();

        public bool IsExpanded
        {
            get => _isExpanded;
            set
            {
                if (_isExpanded != value)
                {
                    _isExpanded = value;
                    OnPropertyChanged();
                }
            }
        }

        public bool IsSelected
        {
            get => _isSelected;
            set
            {
                if (_isSelected != value)
                {
                    _isSelected = value;
                    OnPropertyChanged();
                }
            }
        }

        public bool IsMatching
        {
            get => _isMatching;
            set
            {
                if (_isMatching != value)
                {
                    _isMatching = value;
                    OnPropertyChanged();
                }
            }
        }

        public event PropertyChangedEventHandler? PropertyChanged;

        protected void OnPropertyChanged([CallerMemberName] string? propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        public void ExpandParents()
        {
            var p = Parent;
            while (p != null)
            {
                p.IsExpanded = true;
                p = p.Parent;
            }
        }

        public override string ToString() => DisplayPath;
    }
}
