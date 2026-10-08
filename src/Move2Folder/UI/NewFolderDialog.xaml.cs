using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Move2Folder.Services;

namespace Move2Folder.UI
{
    public partial class NewFolderDialog : Window
    {
        private static readonly char[] InvalidChars = new[] { '\\', '/', ':', '*', '?', '"', '<', '>', '|' };

        private readonly HashSet<string> _existingChildNames;

        public string NewFolderName { get; private set; } = string.Empty;

        public NewFolderDialog(string parentDisplayPath, IEnumerable<string> existingChildNames)
        {
            InitializeComponent();

            _existingChildNames = new HashSet<string>(
                existingChildNames ?? Enumerable.Empty<string>(),
                StringComparer.CurrentCultureIgnoreCase);

            Title = LocalizationService.NewFolderDialogTitle;
            TxtHeaderTitle.Text = LocalizationService.NewFolderHeaderTitle;
            TxtParentPrefix.Text = LocalizationService.NewFolderParentLabel;
            TxtParentPath.Text = parentDisplayPath;
            TxtNameLabel.Text = LocalizationService.NewFolderNameLabel;
            BtnCancel.Content = LocalizationService.BtnCancelShort;
            BtnOk.Content = LocalizationService.BtnCreate;

            string defaultName = GetDefaultNewFolderName(_existingChildNames);
            TxtFolderName.Text = defaultName;

            Loaded += NewFolderDialog_Loaded;
            Activated += (s, e) => LocalizationService.EnforceGermanKeyboardLayout();
        }

        private void NewFolderDialog_Loaded(object sender, RoutedEventArgs e)
        {
            LocalizationService.EnforceGermanKeyboardLayout();
            TxtFolderName.Focus();
            TxtFolderName.SelectAll();
            ValidateInput();
        }

        private static string GetDefaultNewFolderName(HashSet<string> existingNames)
        {
            string baseName = LocalizationService.NewFolderDefaultName;
            if (!existingNames.Contains(baseName))
                return baseName;

            int counter = 2;
            while (existingNames.Contains($"{baseName} ({counter})"))
            {
                counter++;
            }
            return $"{baseName} ({counter})";
        }

        private void TxtFolderName_TextChanged(object sender, TextChangedEventArgs e)
        {
            ValidateInput();
        }

        private void TxtFolderName_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter && BtnOk.IsEnabled)
            {
                BtnOk_Click(sender, e);
                e.Handled = true;
            }
        }

        private void ValidateInput()
        {
            if (TxtFolderName == null || BtnOk == null || TxtError == null)
                return;

            string text = TxtFolderName.Text;

            if (string.IsNullOrWhiteSpace(text))
            {
                BtnOk.IsEnabled = false;
                TxtError.Text = LocalizationService.ErrFolderNameEmpty;
                TxtError.Visibility = Visibility.Visible;
                return;
            }

            string trimmed = text.Trim();

            if (trimmed.IndexOfAny(InvalidChars) >= 0 || trimmed.Any(c => c < 32))
            {
                BtnOk.IsEnabled = false;
                TxtError.Text = LocalizationService.ErrFolderNameInvalidChars;
                TxtError.Visibility = Visibility.Visible;
                return;
            }

            if (trimmed.Length > 255)
            {
                BtnOk.IsEnabled = false;
                TxtError.Text = LocalizationService.ErrFolderNameTooLong;
                TxtError.Visibility = Visibility.Visible;
                return;
            }

            if (_existingChildNames.Contains(trimmed))
            {
                BtnOk.IsEnabled = false;
                TxtError.Text = LocalizationService.ErrFolderAlreadyExists;
                TxtError.Visibility = Visibility.Visible;
                return;
            }

            BtnOk.IsEnabled = true;
            TxtError.Visibility = Visibility.Collapsed;
        }

        private void BtnOk_Click(object sender, RoutedEventArgs e)
        {
            ValidateInput();
            if (!BtnOk.IsEnabled)
                return;

            NewFolderName = TxtFolderName.Text.Trim();
            DialogResult = true;
            Close();
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
