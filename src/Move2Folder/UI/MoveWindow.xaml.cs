using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Move2Folder.Models;
using Move2Folder.Services;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace Move2Folder.UI
{
    public partial class MoveWindow : Window
    {
        private readonly OutlookFolderService _folderService;
        private readonly List<object> _itemsToMove;
        private readonly string? _currentStoreId;
        private readonly string? _suggestedFolderEntryId;
        private List<FolderItem> _flatList = new List<FolderItem>();
        private ObservableCollection<FolderItem> _roots = new ObservableCollection<FolderItem>();
        private FolderItem? _selectedFolder;
        private readonly UserSettings _settings;

        public bool ItemsMoved { get; private set; }

        public MoveWindow(
            OutlookFolderService folderService,
            List<object> itemsToMove,
            string? firstSubject = null,
            string? currentStoreId = null,
            string? suggestedFolderEntryId = null)
        {
            InitializeComponent();
            _folderService = folderService ?? throw new ArgumentNullException(nameof(folderService));
            _itemsToMove = itemsToMove ?? new List<object>();
            _currentStoreId = currentStoreId;
            _suggestedFolderEntryId = suggestedFolderEntryId;
            _settings = SettingsService.Load();

            // Set window size from settings if valid
            if (_settings.WindowWidth >= MinWidth) Width = _settings.WindowWidth;
            if (_settings.WindowHeight >= MinHeight) Height = _settings.WindowHeight;

            // Apply localized text
            Title = LocalizationService.WindowTitle;
            PlaceholderText.Text = LocalizationService.SearchPlaceholder;
            TxtSelectedPathPrefix.Text = LocalizationService.DestinationLabel;
            TxtSelectedPath.Text = LocalizationService.NoFolderSelected;
            TxtShortcutGuide.Text = LocalizationService.ShortcutGuide;
            BtnCancel.Content = LocalizationService.BtnCancel;
            BtnOpenOnly.Content = LocalizationService.BtnOpenOnly;
            BtnOpenOnly.ToolTip = LocalizationService.BtnOpenOnlyTooltip;
            BtnRefresh.ToolTip = LocalizationService.BtnRefreshTooltip;
            BtnNewFolder.ToolTip = LocalizationService.BtnNewFolderTooltip;

            // Context menu items
            MenuNewFolder.Header = LocalizationService.MenuNewFolder;
            MenuNewFolder.InputGestureText = LocalizationService.IsGerman ? "Strg+Umschalt+N" : "Ctrl+Shift+N";
            MenuMoveHere.Header = LocalizationService.MenuMoveHere;
            MenuOpenFolder.Header = LocalizationService.MenuOpenFolder;
            MenuRefresh.Header = LocalizationService.MenuRefresh;

            // Configure header and buttons based on item selection
            int count = _itemsToMove.Count;
            if (count > 0)
            {
                HeaderIcon.Text = "\uE72D"; // Move icon
                TxtItemCount.Text = count == 1 ? LocalizationService.ItemCountSingular : LocalizationService.ItemCountPlural(count);
                TxtItemSubject.Text = string.IsNullOrWhiteSpace(firstSubject) ? LocalizationService.NoSubject : firstSubject;
                BtnMove.Content = LocalizationService.BtnMove;
                BtnOpenOnly.Visibility = Visibility.Visible;
            }
            else
            {
                HeaderIcon.Text = "\uED25"; // Folder icon
                TxtItemCount.Text = LocalizationService.SwitchFolderTitle;
                TxtItemSubject.Text = LocalizationService.SwitchFolderSubtitle;
                BtnMove.Content = LocalizationService.BtnSwitchToFolder;
                BtnOpenOnly.Visibility = Visibility.Collapsed;
            }

            Loaded += MoveWindow_Loaded;
            Activated += MoveWindow_Activated;
            Closing += MoveWindow_Closing;
            SearchTextBox.GotFocus += (s, e) => EnforceGermanKeyboardLayout();
        }

        private void MoveWindow_Loaded(object sender, RoutedEventArgs e)
        {
            EnforceGermanKeyboardLayout();
            LoadFolderData(forceRefresh: false);

            // Pre-select suggested conversation folder or fallback to last used folder
            SelectInitialFolder();

            // Always ensure search box is empty and focused on open
            SearchTextBox.Text = string.Empty;
            PlaceholderText.Visibility = Visibility.Visible;
            SearchTextBox.Focus();
        }

        private void MoveWindow_Activated(object? sender, EventArgs e)
        {
            EnforceGermanKeyboardLayout();
        }

        private void EnforceGermanKeyboardLayout()
        {
            LocalizationService.EnforceGermanKeyboardLayout();
        }

        private void MoveWindow_Closing(object? sender, System.ComponentModel.CancelEventArgs e)
        {
            _settings.WindowWidth = ActualWidth;
            _settings.WindowHeight = ActualHeight;
            SettingsService.Save(_settings);

            // Clean up COM references if items were not moved
            if (!ItemsMoved && _itemsToMove != null)
            {
                foreach (var item in _itemsToMove)
                {
                    try
                    {
                        Marshal.ReleaseComObject(item);
                    }
                    catch { }
                }
            }
        }

        private void LoadFolderData(bool forceRefresh)
        {
            var data = _folderService.LoadFolders(_currentStoreId, forceRefresh);
            _roots = data.Roots;
            _flatList = data.FlatList;
            FolderTree.ItemsSource = _roots;

            // Automatically expand the mailbox root so top-level folders are immediately visible
            if (_roots.Count > 0 && !_roots[0].IsExpanded)
            {
                _roots[0].IsExpanded = true;
            }
        }

        private void SelectInitialFolder()
        {
            // 1. Try to pre-select the uniquely most frequent folder from conversation analysis
            if (!string.IsNullOrEmpty(_suggestedFolderEntryId))
            {
                var target = _flatList.FirstOrDefault(f => !f.IsStore && f.EntryId == _suggestedFolderEntryId);
                if (target != null)
                {
                    SelectFolderInTree(target);
                    return;
                }
            }

            // 2. Fall back to last used folder
            SelectLastUsedFolder();
        }

        private void SelectLastUsedFolder()
        {
            if (string.IsNullOrEmpty(_settings.LastFolderEntryId) && string.IsNullOrEmpty(_settings.LastFolderPath))
                return;

            // Only pre-select if last used folder belongs to the active mailbox
            if (!string.IsNullOrEmpty(_currentStoreId) &&
                !string.IsNullOrEmpty(_settings.LastFolderStoreId) &&
                !string.Equals(_settings.LastFolderStoreId, _currentStoreId, StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            FolderItem? target = null;

            if (!string.IsNullOrEmpty(_settings.LastFolderEntryId))
            {
                target = _flatList.FirstOrDefault(f => f.EntryId == _settings.LastFolderEntryId);
            }

            if (target == null && !string.IsNullOrEmpty(_settings.LastFolderPath))
            {
                target = _flatList.FirstOrDefault(f => string.Equals(f.FullPath, _settings.LastFolderPath, StringComparison.OrdinalIgnoreCase));
            }

            if (target != null)
            {
                SelectFolderInTree(target);
            }
        }

        private void SelectFolderInTree(FolderItem folder)
        {
            folder.ExpandParents();
            folder.IsSelected = true;
            _selectedFolder = folder;
            TxtSelectedPath.Text = folder.DisplayPath;
            BtnMove.IsEnabled = !folder.IsStore;
            BtnOpenOnly.IsEnabled = !folder.IsStore;

            // Bring into view after UI layout updates
            Dispatcher.BeginInvoke(new Action(() =>
            {
                ScrollItemIntoView(FolderTree, folder);
            }), System.Windows.Threading.DispatcherPriority.Background);
        }

        private static bool ScrollItemIntoView(ItemsControl container, FolderItem item)
        {
            foreach (object current in container.Items)
            {
                var treeViewItem = container.ItemContainerGenerator.ContainerFromItem(current) as TreeViewItem;
                if (treeViewItem == null)
                    continue;

                if (current == item)
                {
                    treeViewItem.BringIntoView();
                    return true;
                }

                if (treeViewItem.Items.Count > 0)
                {
                    bool wasExpanded = treeViewItem.IsExpanded;
                    if (!wasExpanded)
                        treeViewItem.IsExpanded = true;

                    treeViewItem.UpdateLayout();
                    if (ScrollItemIntoView(treeViewItem, item))
                        return true;

                    if (!wasExpanded)
                        treeViewItem.IsExpanded = false;
                }
            }
            return false;
        }

        private void SearchTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            string query = SearchTextBox.Text.Trim();
            PlaceholderText.Visibility = string.IsNullOrEmpty(query) ? Visibility.Visible : Visibility.Collapsed;

            if (string.IsNullOrEmpty(query))
            {
                foreach (var f in _flatList)
                {
                    f.IsMatching = false;
                }
                return;
            }

            // Perform type-ahead matching (strictly scoped to current mailbox)
            var matches = _flatList
                .Where(f => !f.IsStore && (string.IsNullOrEmpty(_currentStoreId) || string.Equals(f.StoreId, _currentStoreId, StringComparison.OrdinalIgnoreCase)) && (
                    f.Name.Contains(query, StringComparison.OrdinalIgnoreCase) ||
                    f.DisplayPath.Contains(query, StringComparison.OrdinalIgnoreCase)))
                .OrderBy(f =>
                {
                    if (f.Name.StartsWith(query, StringComparison.OrdinalIgnoreCase)) return 0;
                    if (f.Name.Contains(query, StringComparison.OrdinalIgnoreCase)) return 1;
                    return 2;
                })
                .ThenBy(f => f.Name.Length)
                .ToList();

            // Update matching highlight
            var matchSet = new HashSet<FolderItem>(matches);
            foreach (var f in _flatList)
            {
                f.IsMatching = matchSet.Contains(f);
            }

            // Auto-expand and select first matching folder
            var firstMatch = matches.FirstOrDefault();
            if (firstMatch != null)
            {
                SelectFolderInTree(firstMatch);
            }
        }

        private void SearchTextBox_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Down)
            {
                NavigateRelative(1);
                e.Handled = true;
            }
            else if (e.Key == Key.Up)
            {
                NavigateRelative(-1);
                e.Handled = true;
            }
        }

        private void NavigateRelative(int direction)
        {
            string query = SearchTextBox.Text.Trim();
            List<FolderItem> currentList;

            if (!string.IsNullOrEmpty(query))
            {
                currentList = _flatList.Where(f => f.IsMatching && !f.IsStore && (string.IsNullOrEmpty(_currentStoreId) || string.Equals(f.StoreId, _currentStoreId, StringComparison.OrdinalIgnoreCase))).ToList();
            }
            else
            {
                currentList = _flatList.Where(f => !f.IsStore && (string.IsNullOrEmpty(_currentStoreId) || string.Equals(f.StoreId, _currentStoreId, StringComparison.OrdinalIgnoreCase))).ToList();
            }

            if (currentList.Count == 0)
                return;

            int currentIndex = _selectedFolder != null ? currentList.IndexOf(_selectedFolder) : -1;
            int nextIndex = currentIndex + direction;

            if (nextIndex < 0) nextIndex = 0;
            if (nextIndex >= currentList.Count) nextIndex = currentList.Count - 1;

            SelectFolderInTree(currentList[nextIndex]);
        }

        private void FolderTree_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            if (e.NewValue is FolderItem item)
            {
                _selectedFolder = item;
                TxtSelectedPath.Text = item.DisplayPath;
                BtnMove.IsEnabled = !item.IsStore;
                BtnOpenOnly.IsEnabled = !item.IsStore;
            }
        }

        private void FolderTree_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (_selectedFolder != null && !_selectedFolder.IsStore)
            {
                if (_itemsToMove.Count > 0)
                {
                    ExecuteMove(switchToFolder: false);
                }
                else
                {
                    ExecuteSwitchFolderOnly();
                }
            }
        }

        private void BtnRefresh_Click(object sender, RoutedEventArgs e)
        {
            LoadFolderData(forceRefresh: true);
            SearchTextBox.Text = string.Empty;
            SelectLastUsedFolder();
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        private void BtnOpenOnly_Click(object sender, RoutedEventArgs e)
        {
            ExecuteSwitchFolderOnly();
        }

        private void BtnMove_Click(object sender, RoutedEventArgs e)
        {
            bool isCtrl = (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control;
            if (isCtrl)
            {
                ExecuteMove(switchToFolder: true);
            }
            else if (_itemsToMove.Count > 0)
            {
                ExecuteMove(switchToFolder: false);
            }
            else
            {
                ExecuteSwitchFolderOnly();
            }
        }

        private void Window_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            bool isCtrl = (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control;
            bool isShift = (Keyboard.Modifiers & ModifierKeys.Shift) == ModifierKeys.Shift;

            if (e.Key == Key.Enter)
            {
                if (_selectedFolder != null && !_selectedFolder.IsStore)
                {
                    if (isShift)
                    {
                        // Shift + Enter: Switch to folder without moving
                        ExecuteSwitchFolderOnly();
                        e.Handled = true;
                    }
                    else if (isCtrl)
                    {
                        // Ctrl + Enter: Move emails and switch to folder
                        ExecuteMove(switchToFolder: true);
                        e.Handled = true;
                    }
                    else
                    {
                        // Enter: Move if items selected, else switch to folder
                        if (_itemsToMove.Count > 0)
                        {
                            ExecuteMove(switchToFolder: false);
                        }
                        else
                        {
                            ExecuteSwitchFolderOnly();
                        }
                        e.Handled = true;
                    }
                }
            }
            else if (e.Key == Key.Escape)
            {
                DialogResult = false;
                Close();
                e.Handled = true;
            }
            else if (e.Key == Key.F5)
            {
                BtnRefresh_Click(sender, e);
                e.Handled = true;
            }
            else if (e.Key == Key.N && isCtrl && isShift)
            {
                ShowNewFolderDialog();
                e.Handled = true;
            }
            else if (e.Key == Key.Insert && FolderTree.IsKeyboardFocusWithin)
            {
                ShowNewFolderDialog();
                e.Handled = true;
            }
        }

        private void ExecuteMove(bool switchToFolder)
        {
            if (_selectedFolder == null || _selectedFolder.IsStore)
            {
                MessageBox.Show(
                    LocalizationService.MsgSelectValidFolder,
                    "Move2Folder",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
                return;
            }

            try
            {
                var targetFolder = _folderService.GetFolder(_selectedFolder.EntryId, _selectedFolder.StoreId);
                if (targetFolder == null)
                {
                    MessageBox.Show(
                        LocalizationService.MsgFolderOpenError(_selectedFolder.Name),
                        "Move2Folder",
                        MessageBoxButton.OK,
                        MessageBoxImage.Warning);
                    return;
                }

                if (_itemsToMove.Count > 0)
                {
                    int moved = _folderService.MoveItems(_itemsToMove, targetFolder);
                    ItemsMoved = true;
                }

                // Save selected folder as last used
                _settings.LastFolderEntryId = _selectedFolder.EntryId;
                _settings.LastFolderStoreId = _selectedFolder.StoreId;
                _settings.LastFolderPath = _selectedFolder.FullPath;
                SettingsService.Save(_settings);

                if (switchToFolder)
                {
                    SwitchToTargetFolder(targetFolder);
                }

                Marshal.ReleaseComObject(targetFolder);

                DialogResult = true;
                Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    LocalizationService.MsgMoveError(ex.Message),
                    "Move2Folder",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }

        private void ExecuteSwitchFolderOnly()
        {
            if (_selectedFolder == null || _selectedFolder.IsStore)
            {
                MessageBox.Show(
                    LocalizationService.MsgSelectValidFolder,
                    "Move2Folder",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
                return;
            }

            try
            {
                var targetFolder = _folderService.GetFolder(_selectedFolder.EntryId, _selectedFolder.StoreId);
                if (targetFolder == null)
                {
                    MessageBox.Show(
                        LocalizationService.MsgFolderOpenError(_selectedFolder.Name),
                        "Move2Folder",
                        MessageBoxButton.OK,
                        MessageBoxImage.Warning);
                    return;
                }

                // Save selected folder as last used
                _settings.LastFolderEntryId = _selectedFolder.EntryId;
                _settings.LastFolderStoreId = _selectedFolder.StoreId;
                _settings.LastFolderPath = _selectedFolder.FullPath;
                SettingsService.Save(_settings);

                SwitchToTargetFolder(targetFolder);

                Marshal.ReleaseComObject(targetFolder);

                DialogResult = true;
                Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    LocalizationService.MsgSwitchError(ex.Message),
                    "Move2Folder",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }

        private void SwitchToTargetFolder(Outlook.MAPIFolder targetFolder)
        {
            try
            {
                Outlook.Explorer? explorer = _folderService.OutlookApp.ActiveExplorer();
                if (explorer != null)
                {
                    explorer.CurrentFolder = targetFolder;
                    Marshal.ReleaseComObject(explorer);
                }
                else
                {
                    targetFolder.Display();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error switching to target folder: {ex.Message}");
            }
        }

        private void FolderTree_PreviewMouseRightButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.OriginalSource is DependencyObject source)
            {
                var treeViewItem = FindVisualParent<TreeViewItem>(source);
                if (treeViewItem != null)
                {
                    treeViewItem.Focus();
                    treeViewItem.IsSelected = true;
                    if (treeViewItem.DataContext is FolderItem folder)
                    {
                        _selectedFolder = folder;
                        TxtSelectedPath.Text = folder.DisplayPath;
                        BtnMove.IsEnabled = !folder.IsStore;
                        BtnOpenOnly.IsEnabled = !folder.IsStore;
                    }
                }
            }
        }

        private void FolderContextMenu_Opened(object sender, RoutedEventArgs e)
        {
            if (_selectedFolder == null && _roots.Count > 0)
            {
                SelectFolderInTree(_roots[0]);
            }

            bool hasSelection = _selectedFolder != null;
            bool isStore = _selectedFolder?.IsStore == true;
            bool hasItemsToMove = _itemsToMove.Count > 0;

            MenuNewFolder.IsEnabled = hasSelection;
            MenuMoveHere.Visibility = hasItemsToMove ? Visibility.Visible : Visibility.Collapsed;
            MenuMoveHere.IsEnabled = hasSelection && !isStore;
            MenuOpenFolder.IsEnabled = hasSelection && !isStore;
            MenuSeparator1.Visibility = hasItemsToMove ? Visibility.Visible : Visibility.Collapsed;
        }

        private void MenuNewFolder_Click(object sender, RoutedEventArgs e)
        {
            ShowNewFolderDialog();
        }

        private void MenuMoveHere_Click(object sender, RoutedEventArgs e)
        {
            ExecuteMove(switchToFolder: false);
        }

        private void MenuOpenFolder_Click(object sender, RoutedEventArgs e)
        {
            ExecuteSwitchFolderOnly();
        }

        private void MenuRefresh_Click(object sender, RoutedEventArgs e)
        {
            BtnRefresh_Click(sender, e);
        }

        private void BtnNewFolder_Click(object sender, RoutedEventArgs e)
        {
            ShowNewFolderDialog();
        }

        private void ShowNewFolderDialog()
        {
            if (_selectedFolder == null)
            {
                if (_roots.Count > 0)
                {
                    SelectFolderInTree(_roots[0]);
                }
                else
                {
                    MessageBox.Show(
                        LocalizationService.MsgSelectValidFolder,
                        "Move2Folder",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information);
                    return;
                }
            }

            var parentFolder = _selectedFolder!;
            var dialog = new NewFolderDialog(parentFolder.DisplayPath, parentFolder.Children.Select(c => c.Name))
            {
                Owner = this
            };

            if (dialog.ShowDialog() == true)
            {
                string newFolderName = dialog.NewFolderName;
                try
                {
                    var createdItem = _folderService.CreateFolder(parentFolder, newFolderName);
                    if (createdItem != null)
                    {
                        if (!_flatList.Contains(createdItem))
                        {
                            _flatList.Add(createdItem);
                        }

                        // Clear search filter so new folder is visible in its full tree path
                        if (!string.IsNullOrEmpty(SearchTextBox.Text))
                        {
                            SearchTextBox.Text = string.Empty;
                        }

                        // Expand parent and select the newly created folder
                        SelectFolderInTree(createdItem);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(
                        LocalizationService.MsgCreateFolderError(newFolderName, ex.Message),
                        "Move2Folder",
                        MessageBoxButton.OK,
                        MessageBoxImage.Error);
                }
            }
        }

        private static T? FindVisualParent<T>(DependencyObject child) where T : DependencyObject
        {
            var current = child;
            while (current != null)
            {
                if (current is T match)
                    return match;
                current = System.Windows.Media.VisualTreeHelper.GetParent(current);
            }
            return null;
        }
    }
}
