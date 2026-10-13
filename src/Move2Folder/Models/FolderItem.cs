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
        private string _name = string.Empty;
        private string _fullPath = string.Empty;
        private string _displayPath = string.Empty;
        private string _folderType = "Folder";

        public string EntryId { get; set; } = string.Empty;
        public string StoreId { get; set; } = string.Empty;

        public string Name
        {
            get => _name;
            set
            {
                if (_name != value)
                {
                    _name = value;
                    OnPropertyChanged();
                }
            }
        }

        public string FullPath
        {
            get => _fullPath;
            set
            {
                if (_fullPath != value)
                {
                    _fullPath = value;
                    OnPropertyChanged();
                }
            }
        }

        public string DisplayPath
        {
            get => _displayPath;
            set
            {
                if (_displayPath != value)
                {
                    _displayPath = value;
                    OnPropertyChanged();
                }
            }
        }

        public bool IsStore { get; set; }

        public string FolderType
        {
            get => _folderType;
            set
            {
                if (_folderType != value)
                {
                    _folderType = value;
                    OnPropertyChanged();
                }
            }
        }

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
