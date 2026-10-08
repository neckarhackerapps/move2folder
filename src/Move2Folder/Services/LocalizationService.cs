using System;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace Move2Folder.Services
{
    public static class LocalizationService
    {
        private const int MsoLanguageIDUI = 2; // Microsoft.Office.Core.MsoAppLanguageID.msoLanguageIDUI

        [DllImport("user32.dll")]
        private static extern uint GetKeyboardLayoutList(int nBuff, [Out] IntPtr[]? lpList);

        [DllImport("user32.dll")]
        private static extern IntPtr ActivateKeyboardLayout(IntPtr hkl, uint flags);

        public static bool IsGerman { get; private set; } = true;

        public static void Initialize(Outlook.Application? outlookApp)
        {
            try
            {
                if (outlookApp != null)
                {
                    dynamic languageSettings = outlookApp.LanguageSettings;
                    int lcid = languageSettings.LanguageID(MsoLanguageIDUI);
                    var culture = CultureInfo.GetCultureInfo(lcid);
                    IsGerman = culture.TwoLetterISOLanguageName.Equals("de", StringComparison.OrdinalIgnoreCase);
                    return;
                }
            }
            catch
            {
                // Fallback to CurrentUICulture
            }

            IsGerman = CultureInfo.CurrentUICulture.TwoLetterISOLanguageName.Equals("de", StringComparison.OrdinalIgnoreCase);
        }

        public static void EnforceGermanKeyboardLayout()
        {
            // Only enforce German layout if Outlook/system language is German!
            if (!IsGerman)
                return;

            try
            {
                // 1. Force WPF input language manager to German
                System.Windows.Input.InputLanguageManager.Current.CurrentInputLanguage =
                    new CultureInfo("de-DE");

                // 2. Force Win32 thread keyboard layout to German (LangID 0x0407)
                uint count = GetKeyboardLayoutList(0, null);
                if (count > 0)
                {
                    IntPtr[] list = new IntPtr[count];
                    GetKeyboardLayoutList((int)count, list);
                    IntPtr german = list.FirstOrDefault(l => (l.ToInt64() & 0xFFFF) == 0x0407);
                    if (german != IntPtr.Zero)
                    {
                        ActivateKeyboardLayout(german, 0);
                    }
                }
            }
            catch { }
        }

        // Window & Header
        public static string WindowTitle =>
            IsGerman ? "Move2Folder - Zielordner wählen" : "Move2Folder - Select Target Folder";

        public static string SearchPlaceholder =>
            IsGerman ? "Zielordner suchen" : "Search target folder";

        public static string ItemCountSingular =>
            IsGerman ? "1 E-Mail ausgewählt" : "1 email selected";

        public static string ItemCountPlural(int count) =>
            IsGerman ? $"{count} E-Mails ausgewählt" : $"{count} emails selected";

        public static string SwitchFolderTitle =>
            IsGerman ? "Ordner wechseln" : "Switch Folder";

        public static string SwitchFolderSubtitle =>
            IsGerman
                ? "Keine E-Mail ausgewählt – Zielordner ansteuern und Enter drücken"
                : "No email selected – select destination folder and press Enter";

        public static string NoSubject =>
            IsGerman ? "(Kein Betreff)" : "(No subject)";

        public static string UntitledFolder =>
            IsGerman ? "Unbenannt" : "Untitled";

        // Target Path Display
        public static string DestinationLabel =>
            IsGerman ? "Ziel: " : "Target: ";

        public static string NoFolderSelected =>
            IsGerman ? "Kein Ordner ausgewählt" : "No folder selected";

        // Shortcut Guide
        public static string ShortcutGuide =>
            IsGerman
                ? "↵ Enter: Verschieben  |  Strg+Enter: Verschieben & Öffnen  |  Shift+Enter: Nur Öffnen  |  Esc: Abbrechen"
                : "↵ Enter: Move  |  Ctrl+Enter: Move & Open  |  Shift+Enter: Open Only  |  Esc: Cancel";

        // Buttons
        public static string BtnCancel =>
            IsGerman ? "Abbrechen (Esc)" : "Cancel (Esc)";

        public static string BtnCancelShort =>
            IsGerman ? "Abbrechen" : "Cancel";

        public static string BtnCreate =>
            IsGerman ? "Erstellen" : "Create";

        public static string BtnOpenOnly =>
            IsGerman ? "Nur Öffnen (Shift+↵)" : "Open Only (Shift+↵)";

        public static string BtnOpenOnlyTooltip =>
            IsGerman
                ? "Wechselt in den Zielordner, ohne E-Mails zu verschieben (Shift+Enter)"
                : "Switches to destination folder without moving emails (Shift+Enter)";

        public static string BtnMove =>
            IsGerman ? "Verschieben (Enter)" : "Move (Enter)";

        public static string BtnSwitchToFolder =>
            IsGerman ? "In Ordner wechseln (Enter)" : "Switch to Folder (Enter)";

        public static string BtnRefreshTooltip =>
            IsGerman ? "Ordnerliste neu laden (F5)" : "Refresh folder list (F5)";

        public static string BtnNewFolderTooltip =>
            IsGerman ? "Neuen Ordner erstellen (Strg+Umschalt+N)" : "Create new folder (Ctrl+Shift+N)";

        // Context Menu
        public static string MenuNewFolder =>
            IsGerman ? "Neuer Ordner..." : "New Folder...";

        public static string MenuMoveHere =>
            IsGerman ? "In diesen Ordner verschieben" : "Move to this folder";

        public static string MenuOpenFolder =>
            IsGerman ? "Ordner öffnen" : "Open folder";

        public static string MenuRefresh =>
            IsGerman ? "Aktualisieren" : "Refresh";

        // New Folder Dialog
        public static string NewFolderDialogTitle =>
            IsGerman ? "Neuer Ordner" : "New Folder";

        public static string NewFolderHeaderTitle =>
            IsGerman ? "Neuen Ordner erstellen" : "Create New Folder";

        public static string NewFolderParentLabel =>
            IsGerman ? "Erstellen in: " : "Create in: ";

        public static string NewFolderNameLabel =>
            IsGerman ? "Name des neuen Ordners:" : "New folder name:";

        public static string NewFolderDefaultName =>
            IsGerman ? "Neuer Ordner" : "New Folder";

        public static string ErrFolderNameEmpty =>
            IsGerman ? "Der Ordnername darf nicht leer sein." : "Folder name cannot be empty.";

        public static string ErrFolderNameInvalidChars =>
            IsGerman
                ? "Der Name darf keines der folgenden Zeichen enthalten: \\ / : * ? \" < > |"
                : "The name cannot contain any of the following characters: \\ / : * ? \" < > |";

        public static string ErrFolderNameTooLong =>
            IsGerman ? "Der Ordnername ist zu lang (maximal 255 Zeichen)." : "Folder name is too long (maximum 255 characters).";

        public static string ErrFolderAlreadyExists =>
            IsGerman ? "Ein Ordner mit diesem Namen existiert hier bereits." : "A folder with this name already exists in this location.";

        // Messages & Errors
        public static string MsgSelectValidFolder =>
            IsGerman
                ? "Bitte wählen Sie einen gültigen Zielordner aus."
                : "Please select a valid destination folder.";

        public static string MsgFolderOpenError(string folderName) =>
            IsGerman
                ? $"Der Zielordner '{folderName}' konnte in Outlook nicht geöffnet werden."
                : $"The target folder '{folderName}' could not be opened in Outlook.";

        public static string MsgCreateFolderError(string folderName, string message) =>
            IsGerman
                ? $"Der Ordner '{folderName}' konnte nicht erstellt werden:\n{message}"
                : $"Could not create folder '{folderName}':\n{message}";

        public static string MsgMoveError(string message) =>
            IsGerman
                ? $"Fehler beim Verschieben: {message}"
                : $"Error moving item(s): {message}";

        public static string MsgSwitchError(string message) =>
            IsGerman
                ? $"Fehler beim Wechseln in den Ordner: {message}"
                : $"Error switching to folder: {message}";

        public static string MsgUnexpectedError(string message) =>
            IsGerman
                ? $"Unerwarteter Fehler in Move2Folder: {message}"
                : $"Unexpected error in Move2Folder: {message}";

        // Ribbon
        public static string RibbonButtonLabel =>
            IsGerman ? "In Ordner verschieben" : "Move to Folder";

        public static string RibbonScreentip =>
            IsGerman ? "Move2Folder (Strg+Umschalt+V)" : "Move2Folder (Ctrl+Shift+V)";

        public static string RibbonSupertip =>
            IsGerman
                ? "Verschiebt die ausgewählte(n) E-Mail(s) in einen Zielordner des aktuellen Postfachs mit schneller Suche (Strg+Shift+V)."
                : "Moves selected email(s) to a target folder in the active mailbox with fast type-ahead search (Ctrl+Shift+V).";
    }
}
