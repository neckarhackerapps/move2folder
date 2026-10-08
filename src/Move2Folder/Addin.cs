using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;
using Move2Folder.Services;
using Move2Folder.UI;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace Move2Folder
{
    [ComVisible(true)]
    [Guid("298621CD-332D-4DE0-B3B8-43BA73CA4893")]
    [ProgId("Move2Folder.Addin")]
    public class Addin : IDTExtensibility2, IRibbonExtensibility
    {
        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        private static void Log(string message)
        {
            try
            {
                string logFile = Path.Combine(Path.GetTempPath(), "move2folder.log");
                File.AppendAllText(logFile, $"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff} [PID {Environment.ProcessId}] {message}\r\n");
            }
            catch { }
        }

        private Outlook.Application? _outlookApp;
        private OutlookFolderService? _folderService;
        private KeyboardHook? _keyboardHook;
        private IRibbonUI? _ribbon;
        private bool _isDialogOpen;

        public Addin()
        {
            Log("Addin initialized.");
        }

        #region IDTExtensibility2 Implementation

        public void OnConnection(
            object Application,
            ext_ConnectMode ConnectMode,
            object AddInInst,
            ref Array custom)
        {
            try
            {
                Log($"OnConnection called with mode {ConnectMode}");
                _outlookApp = (Outlook.Application)Application;
                LocalizationService.Initialize(_outlookApp);
                Log($"Localization initialized: IsGerman={LocalizationService.IsGerman}");

                _folderService = new OutlookFolderService(_outlookApp);

                // Initialize keyboard hook for Ctrl+Shift+V
                _keyboardHook = new KeyboardHook(OnMove2FolderTriggered);
                Log("Keyboard hook installed successfully.");
            }
            catch (Exception ex)
            {
                Log($"Error during Addin OnConnection: {ex}");
                System.Diagnostics.Debug.WriteLine($"Error during Addin OnConnection: {ex}");
            }
        }

        public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            try
            {
                _keyboardHook?.Dispose();
                _keyboardHook = null;
                _folderService = null;

                if (_outlookApp != null)
                {
                    Marshal.ReleaseComObject(_outlookApp);
                    _outlookApp = null;
                }
            }
            catch { }
        }

        public void OnAddInsUpdate(ref Array custom) { }

        public void OnStartupComplete(ref Array custom) { }

        public void OnBeginShutdown(ref Array custom) { }

        #endregion

        #region IRibbonExtensibility Implementation

        public string GetCustomUI(string RibbonID)
        {
            try
            {
                LocalizationService.Initialize(_outlookApp);
                var assembly = Assembly.GetExecutingAssembly();
                using var stream = assembly.GetManifestResourceStream("Move2Folder.Resources.Ribbon.xml");
                if (stream != null)
                {
                    using var reader = new StreamReader(stream);
                    string xml = reader.ReadToEnd();
                    xml = xml.Replace("{RibbonLabel}", LocalizationService.RibbonButtonLabel)
                             .Replace("{RibbonScreentip}", LocalizationService.RibbonScreentip)
                             .Replace("{RibbonSupertip}", LocalizationService.RibbonSupertip);
                    return xml;
                }
            }
            catch (Exception ex)
            {
                Log($"Error loading Ribbon XML: {ex}");
            }

            return string.Empty;
        }

        public void OnRibbonLoad(IRibbonUI ribbon)
        {
            _ribbon = ribbon;
        }

        public void OnMove2FolderRibbonClick(IRibbonControl control)
        {
            OnMove2FolderTriggered();
        }

        #endregion

        #region Move Action & UI Dialog

        public void OnMove2FolderTriggered()
        {
            if (_outlookApp == null || _folderService == null)
                return;

            // Avoid opening multiple dialogs simultaneously
            if (_isDialogOpen)
                return;

            try
            {
                _isDialogOpen = true;

                var itemsToMove = new List<object>();
                var conversations = new List<Outlook.Conversation>();
                var seenConvIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                void TryAddConversation(Outlook.Conversation? conv)
                {
                    if (conv == null) return;
                    try
                    {
                        string? convId = conv.ConversationID;
                        if (!string.IsNullOrEmpty(convId) && seenConvIds.Add(convId))
                        {
                            conversations.Add(conv);
                            return;
                        }
                    }
                    catch { }
                    try { Marshal.ReleaseComObject(conv); } catch { }
                }

                string? firstSubject = null;
                string? currentStoreId = null;
                Outlook.Inspector? inspectorToClose = null;

                // Determine whether the active window is an Inspector or an Explorer
                object? activeWindow = null;
                try
                {
                    activeWindow = _outlookApp.ActiveWindow();
                }
                catch (Exception ex)
                {
                    Log($"Error obtaining ActiveWindow: {ex.Message}");
                }

                if (activeWindow is Outlook.Inspector activeInsp)
                {
                    try
                    {
                        object? currentItem = activeInsp.CurrentItem;
                        if (currentItem != null)
                        {
                            if (IsDraft(currentItem))
                            {
                                Log("Active window is an Inspector containing a draft. Drafts are protected and will not be moved.");
                                Marshal.ReleaseComObject(currentItem);
                                Marshal.ReleaseComObject(activeInsp);
                                return;
                            }

                            itemsToMove.Add(currentItem);
                            firstSubject = GetItemSubject(currentItem);
                            currentStoreId = GetItemStoreId(currentItem);
                            inspectorToClose = activeInsp;

                            var itemConv = GetItemConversation(currentItem);
                            TryAddConversation(itemConv);
                        }
                        else
                        {
                            Marshal.ReleaseComObject(activeInsp);
                        }
                    }
                    catch (Exception ex)
                    {
                        Log($"Error processing active inspector item: {ex.Message}");
                        if (inspectorToClose == null)
                            Marshal.ReleaseComObject(activeInsp);
                    }
                }
                else
                {
                    // Active window is an Explorer (or fallback)
                    Outlook.Explorer? explorer = activeWindow as Outlook.Explorer ?? _outlookApp.ActiveExplorer();
                    if (explorer != null)
                    {
                        try
                        {
                            Outlook.Selection? selection = explorer.Selection;
                            if (selection != null)
                            {
                                CollectSelectedItems(selection, itemsToMove, ref firstSubject);
                                Marshal.ReleaseComObject(selection);
                            }

                            try
                            {
                                Outlook.MAPIFolder? curFolder = explorer.CurrentFolder;
                                if (curFolder != null)
                                {
                                    currentStoreId = curFolder.StoreID;
                                    Marshal.ReleaseComObject(curFolder);
                                }
                            }
                            catch { }
                        }
                        finally
                        {
                            Marshal.ReleaseComObject(explorer);
                        }
                    }
                }

                // Discover conversations from all items to move (deduplicated by ConversationID)
                foreach (var item in itemsToMove)
                {
                    string? itemConvId = GetItemConversationId(item);
                    if (!string.IsNullOrEmpty(itemConvId) && seenConvIds.Contains(itemConvId))
                    {
                        continue;
                    }

                    var itemConv = GetItemConversation(item);
                    TryAddConversation(itemConv);
                }

                if (string.IsNullOrEmpty(currentStoreId) && itemsToMove.Count > 0)
                {
                    foreach (var item in itemsToMove)
                    {
                        currentStoreId = GetItemStoreId(item);
                        if (!string.IsNullOrEmpty(currentStoreId))
                            break;
                    }
                }

                if (string.IsNullOrEmpty(currentStoreId))
                {
                    try
                    {
                        var session = _outlookApp.Session;
                        if (session != null)
                        {
                            var defStore = session.DefaultStore;
                            if (defStore != null)
                            {
                                currentStoreId = defStore.StoreID;
                                Marshal.ReleaseComObject(defStore);
                            }
                            Marshal.ReleaseComObject(session);
                        }
                    }
                    catch { }
                }

                // Analyze conversation(s) to determine uniquely most frequent destination folder
                string? suggestedFolderEntryId = DetermineSuggestedConversationFolder(conversations, currentStoreId);

                Log($"Opening MoveWindow with {itemsToMove.Count} item(s) in store '{currentStoreId}'. Suggested folder: '{suggestedFolderEntryId ?? "none"}'.");

                var window = new MoveWindow(_folderService, itemsToMove, firstSubject, currentStoreId, suggestedFolderEntryId);

                // Set foreground window as WPF window owner
                IntPtr activeHwnd = GetForegroundWindow();
                if (activeHwnd != IntPtr.Zero)
                {
                    new WindowInteropHelper(window).Owner = activeHwnd;
                }

                bool? result = window.ShowDialog();

                // If items were moved from an open inspector, close that specific inspector
                if (result == true && window.ItemsMoved && inspectorToClose != null)
                {
                    try
                    {
                        Log("Closing inspector for moved item.");
                        inspectorToClose.Close(Outlook.OlInspectorClose.olDiscard);
                    }
                    catch (Exception ex)
                    {
                        Log($"Error closing inspector: {ex.Message}");
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(inspectorToClose);
                        inspectorToClose = null;
                    }
                }
                else if (inspectorToClose != null)
                {
                    Marshal.ReleaseComObject(inspectorToClose);
                    inspectorToClose = null;
                }
            }
            catch (Exception ex)
            {
                Log($"Unexpected error in OnMove2FolderTriggered: {ex}");
                MessageBox.Show(
                    LocalizationService.MsgUnexpectedError(ex.Message),
                    "Move2Folder",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
            finally
            {
                _isDialogOpen = false;
            }
        }

        private void CollectSelectedItems(
            Outlook.Selection selection,
            List<object> itemsToMove,
            ref string? firstSubject)
        {
            var seenEntryIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            // 1. Check for selected ConversationHeaders via GetSelection(olConversationHeaders)
            Outlook.Selection? convHeaders = null;
            try
            {
                convHeaders = selection.GetSelection(Outlook.OlSelectionContents.olConversationHeaders);
            }
            catch (Exception ex)
            {
                Log($"GetSelection(olConversationHeaders) info: {ex.Message}");
            }

            if (convHeaders != null)
            {
                try
                {
                    int convCount = convHeaders.Count;
                    if (convCount > 0)
                    {
                        Log($"Found {convCount} conversation header(s) selected.");
                        for (int i = 1; i <= convCount; i++)
                        {
                            object headerObj = convHeaders[i];
                            if (headerObj is Outlook.ConversationHeader convHeader)
                            {
                                try
                                {
                                    if (firstSubject == null)
                                    {
                                        firstSubject = convHeader.ConversationTopic;
                                    }

                                    // 1a. Items in current folder belonging to this conversation
                                    Outlook.SimpleItems? simpleItems = null;
                                    try
                                    {
                                        simpleItems = convHeader.GetItems();
                                        if (simpleItems != null)
                                        {
                                            int sCount = simpleItems.Count;
                                            Log($"Conversation '{convHeader.ConversationTopic}' has {sCount} item(s) in folder.");
                                            for (int j = 1; j <= sCount; j++)
                                            {
                                                object sItem = simpleItems[j];
                                                if (!AddItemIfUnique(sItem, itemsToMove, seenEntryIds, ref firstSubject))
                                                {
                                                    Marshal.ReleaseComObject(sItem);
                                                }
                                            }
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        Log($"Error getting items from convHeader: {ex.Message}");
                                    }
                                    finally
                                    {
                                        if (simpleItems != null)
                                            Marshal.ReleaseComObject(simpleItems);
                                    }

                                    // 1b. Also retrieve items from the full Conversation across the mailbox
                                    Outlook.Conversation? conv = null;
                                    try
                                    {
                                        conv = convHeader.GetConversation();
                                        if (conv != null)
                                        {
                                            Outlook.SimpleItems? rootItems = conv.GetRootItems();
                                            if (rootItems != null)
                                            {
                                                int rCount = rootItems.Count;
                                                for (int r = 1; r <= rCount; r++)
                                                {
                                                    object rootItem = rootItems[r];
                                                    try
                                                    {
                                                        bool added = AddItemIfUnique(rootItem, itemsToMove, seenEntryIds, ref firstSubject);
                                                        EnumerateConversationChildren(conv, rootItem, itemsToMove, seenEntryIds, ref firstSubject);
                                                        if (!added)
                                                        {
                                                            Marshal.ReleaseComObject(rootItem);
                                                        }
                                                    }
                                                    catch (Exception ex)
                                                    {
                                                        Log($"Error processing root item: {ex.Message}");
                                                    }
                                                }
                                                Marshal.ReleaseComObject(rootItems);
                                            }
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        Log($"Error traversing Conversation: {ex.Message}");
                                    }
                                    finally
                                    {
                                        if (conv != null)
                                            Marshal.ReleaseComObject(conv);
                                    }
                                }
                                finally
                                {
                                    Marshal.ReleaseComObject(convHeader);
                                }
                            }
                            else
                            {
                                Marshal.ReleaseComObject(headerObj);
                            }
                        }
                    }
                }
                finally
                {
                    Marshal.ReleaseComObject(convHeaders);
                }
            }

            // 2. Also check standard items in selection (or when no conversation headers were selected)
            int count = selection.Count;
            for (int i = 1; i <= count; i++)
            {
                object item = selection[i];
                if (item is Outlook.ConversationHeader ch)
                {
                    try
                    {
                        if (firstSubject == null)
                        {
                            firstSubject = ch.ConversationTopic;
                        }

                        Outlook.SimpleItems? sItems = ch.GetItems();
                        if (sItems != null)
                        {
                            int sCount = sItems.Count;
                            for (int j = 1; j <= sCount; j++)
                            {
                                object sItem = sItems[j];
                                if (!AddItemIfUnique(sItem, itemsToMove, seenEntryIds, ref firstSubject))
                                {
                                    Marshal.ReleaseComObject(sItem);
                                }
                            }
                            Marshal.ReleaseComObject(sItems);
                        }
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(ch);
                    }
                }
                else
                {
                    if (!AddItemIfUnique(item, itemsToMove, seenEntryIds, ref firstSubject))
                    {
                        Marshal.ReleaseComObject(item);
                    }
                }
            }

            Log($"Total items collected to move: {itemsToMove.Count}. First subject: '{firstSubject}'");
        }

        private void EnumerateConversationChildren(
            Outlook.Conversation conv,
            object parentItem,
            List<object> itemsToMove,
            HashSet<string> seenEntryIds,
            ref string? firstSubject)
        {
            Outlook.SimpleItems? children = null;
            try
            {
                children = conv.GetChildren(parentItem);
                if (children != null)
                {
                    int childCount = children.Count;
                    for (int i = 1; i <= childCount; i++)
                    {
                        object child = children[i];
                        try
                        {
                            bool added = AddItemIfUnique(child, itemsToMove, seenEntryIds, ref firstSubject);
                            EnumerateConversationChildren(conv, child, itemsToMove, seenEntryIds, ref firstSubject);
                            if (!added)
                            {
                                Marshal.ReleaseComObject(child);
                            }
                        }
                        catch (Exception ex)
                        {
                            Log($"Error processing child item: {ex.Message}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Log($"EnumerateConversationChildren error: {ex.Message}");
            }
            finally
            {
                if (children != null)
                    Marshal.ReleaseComObject(children);
            }
        }

        private bool AddItemIfUnique(
            object item,
            List<object> itemsToMove,
            HashSet<string> seenEntryIds,
            ref string? firstSubject)
        {
            if (item == null) return false;

            // Drafts must NEVER be moved
            if (IsDraft(item))
            {
                Log("Skipping draft item: drafts are preserved in Drafts folder.");
                return false;
            }

            string? entryId = GetItemEntryId(item);
            if (!string.IsNullOrEmpty(entryId))
            {
                if (!seenEntryIds.Add(entryId))
                {
                    // Already added, skip
                    return false;
                }
            }

            if (firstSubject == null)
            {
                firstSubject = GetItemSubject(item);
            }

            itemsToMove.Add(item);
            return true;
        }

        private string? DetermineSuggestedConversationFolder(
            List<Outlook.Conversation> conversations,
            string? currentStoreId)
        {
            try
            {
                if (conversations.Count == 0)
                {
                    Log("No conversations found for frequency analysis.");
                    return null;
                }

                // Build set of excluded folder EntryIDs (Inbox, Sent, Drafts, Trash, Junk, Outbox, and store roots)
                var excludedFolderEntryIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                void AddExcludedFolder(Outlook.MAPIFolder? folder)
                {
                    if (folder == null) return;
                    try
                    {
                        string eid = folder.EntryID;
                        if (!string.IsNullOrEmpty(eid))
                            excludedFolderEntryIds.Add(eid);
                    }
                    catch { }
                    finally
                    {
                        Marshal.ReleaseComObject(folder);
                    }
                }

                try
                {
                    var session = _outlookApp?.Session;
                    if (session != null)
                    {
                        try
                        {
                            AddExcludedFolder(session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox));
                            AddExcludedFolder(session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderSentMail));
                            AddExcludedFolder(session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderDrafts));
                            AddExcludedFolder(session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderDeletedItems));
                            AddExcludedFolder(session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderJunk));
                            AddExcludedFolder(session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderOutbox));

                            if (!string.IsNullOrEmpty(currentStoreId))
                            {
                                Outlook.Store? store = null;
                                try
                                {
                                    store = session.GetStoreFromID(currentStoreId);
                                    if (store != null)
                                    {
                                        AddExcludedFolder(store.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox));
                                        AddExcludedFolder(store.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderSentMail));
                                        AddExcludedFolder(store.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderDrafts));
                                        AddExcludedFolder(store.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderDeletedItems));
                                        AddExcludedFolder(store.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderJunk));
                                        AddExcludedFolder(store.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderOutbox));
                                        AddExcludedFolder(store.GetRootFolder());
                                    }
                                }
                                catch { }
                                finally
                                {
                                    if (store != null)
                                        Marshal.ReleaseComObject(store);
                                }
                            }
                        }
                        finally
                        {
                            Marshal.ReleaseComObject(session);
                        }
                    }
                }
                catch (Exception ex)
                {
                    Log($"Error collecting default excluded folders: {ex.Message}");
                }

                var folderCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

                foreach (var conv in conversations)
                {
                    try
                    {
                        Outlook.SimpleItems? rootItems = conv.GetRootItems();
                        if (rootItems != null)
                        {
                            try
                            {
                                int rCount = rootItems.Count;
                                for (int r = 1; r <= rCount; r++)
                                {
                                    object rootItem = rootItems[r];
                                    try
                                    {
                                        TallyConversationItem(conv, rootItem, folderCounts, excludedFolderEntryIds, currentStoreId);
                                    }
                                    finally
                                    {
                                        Marshal.ReleaseComObject(rootItem);
                                    }
                                }
                            }
                            finally
                            {
                                Marshal.ReleaseComObject(rootItems);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Log($"Error traversing conversation root items for frequency: {ex.Message}");
                    }
                }

                if (folderCounts.Count == 0)
                {
                    Log("No non-excluded target folders found in conversation items.");
                    return null;
                }

                int maxCount = 0;
                foreach (var count in folderCounts.Values)
                {
                    if (count > maxCount) maxCount = count;
                }

                if (maxCount <= 0)
                    return null;

                var topFolders = new List<string>();
                foreach (var kvp in folderCounts)
                {
                    if (kvp.Value == maxCount)
                    {
                        topFolders.Add(kvp.Key);
                    }
                }

                if (topFolders.Count == 1)
                {
                    string suggestedEntryId = topFolders[0];
                    Log($"Unique most frequent conversation folder identified: EntryID={suggestedEntryId} (frequency={maxCount}).");
                    return suggestedEntryId;
                }

                Log($"Conversation folder frequency tie with count {maxCount} among {topFolders.Count} folders. Fallback to last used folder.");
                return null;
            }
            catch (Exception ex)
            {
                Log($"Error in DetermineSuggestedConversationFolder: {ex}");
                return null;
            }
            finally
            {
                foreach (var conv in conversations)
                {
                    try { Marshal.ReleaseComObject(conv); } catch { }
                }
                conversations.Clear();
            }
        }

        private void TallyConversationItem(
            Outlook.Conversation conv,
            object item,
            Dictionary<string, int> folderCounts,
            HashSet<string> excludedFolderEntryIds,
            string? currentStoreId)
        {
            if (item == null) return;

            try
            {
                if (!IsDraft(item))
                {
                    Outlook.MAPIFolder? parentFolder = null;
                    try
                    {
                        dynamic dyn = item;
                        parentFolder = dyn.Parent as Outlook.MAPIFolder;
                        if (parentFolder != null)
                        {
                            bool storeMatches = string.IsNullOrEmpty(currentStoreId) ||
                                               string.Equals(parentFolder.StoreID, currentStoreId, StringComparison.OrdinalIgnoreCase);

                            if (storeMatches && !IsExcludedFolder(parentFolder, excludedFolderEntryIds))
                            {
                                string eid = parentFolder.EntryID;
                                if (!string.IsNullOrEmpty(eid))
                                {
                                    folderCounts[eid] = folderCounts.GetValueOrDefault(eid) + 1;
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Log($"Error getting parent folder in tally: {ex.Message}");
                    }
                    finally
                    {
                        if (parentFolder != null)
                            Marshal.ReleaseComObject(parentFolder);
                    }
                }

                // Recursively inspect child items in the conversation
                Outlook.SimpleItems? children = null;
                try
                {
                    children = conv.GetChildren(item);
                    if (children != null)
                    {
                        int childCount = children.Count;
                        for (int i = 1; i <= childCount; i++)
                        {
                            object child = children[i];
                            try
                            {
                                TallyConversationItem(conv, child, folderCounts, excludedFolderEntryIds, currentStoreId);
                            }
                            finally
                            {
                                Marshal.ReleaseComObject(child);
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Log($"Error enumerating children in tally: {ex.Message}");
                }
                finally
                {
                    if (children != null)
                        Marshal.ReleaseComObject(children);
                }
            }
            catch (Exception ex)
            {
                Log($"Error in TallyConversationItem: {ex.Message}");
            }
        }

        private bool IsDraft(object item)
        {
            if (item == null) return false;

            try
            {
                // 1. Check Sent property
                if (item is Outlook._MailItem mail)
                {
                    if (!mail.Sent) return true;
                }
                else
                {
                    try
                    {
                        dynamic dyn = item;
                        if (dyn.Sent != null && !(bool)dyn.Sent)
                        {
                            return true;
                        }
                    }
                    catch { }
                }

                // 2. Check parent folder (Drafts / Entwürfe)
                Outlook.MAPIFolder? parent = null;
                try
                {
                    dynamic dyn = item;
                    parent = dyn.Parent as Outlook.MAPIFolder;
                    if (parent != null)
                    {
                        if (IsDraftsFolder(parent))
                        {
                            return true;
                        }
                    }
                }
                catch { }
                finally
                {
                    if (parent != null)
                        Marshal.ReleaseComObject(parent);
                }
            }
            catch (Exception ex)
            {
                Log($"IsDraft error: {ex.Message}");
            }

            return false;
        }

        private bool IsDraftsFolder(Outlook.MAPIFolder folder)
        {
            if (folder == null) return false;

            try
            {
                string name = folder.Name ?? string.Empty;
                if (name.Equals("Drafts", StringComparison.OrdinalIgnoreCase) ||
                    name.Equals("Entwürfe", StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }

                string fullPath = folder.FolderPath ?? string.Empty;
                if (fullPath.EndsWith("\\Drafts", StringComparison.OrdinalIgnoreCase) ||
                    fullPath.EndsWith("\\Entwürfe", StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }

                if (_outlookApp != null)
                {
                    Outlook.NameSpace? session = null;
                    Outlook.MAPIFolder? defaultDrafts = null;
                    try
                    {
                        session = _outlookApp.Session;
                        if (session != null)
                        {
                            defaultDrafts = session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderDrafts);
                            if (defaultDrafts != null && string.Equals(folder.EntryID, defaultDrafts.EntryID, StringComparison.OrdinalIgnoreCase))
                            {
                                return true;
                            }
                        }
                    }
                    catch { }
                    finally
                    {
                        if (defaultDrafts != null) Marshal.ReleaseComObject(defaultDrafts);
                        if (session != null) Marshal.ReleaseComObject(session);
                    }
                }
            }
            catch { }

            return false;
        }

        private static bool IsExcludedFolder(Outlook.MAPIFolder folder, HashSet<string> excludedFolderEntryIds)
        {
            try
            {
                string eid = folder.EntryID;
                if (!string.IsNullOrEmpty(eid) && excludedFolderEntryIds.Contains(eid))
                    return true;

                string name = folder.Name ?? string.Empty;
                if (IsExcludedFolderName(name))
                    return true;
            }
            catch { }
            return false;
        }

        private static bool IsExcludedFolderName(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return true;
            string n = name.Trim().ToLowerInvariant();
            return n == "posteingang" || n == "inbox" ||
                   n == "gesendete elemente" || n == "gesendet" || n == "sent items" || n == "sent" ||
                   n == "entwürfe" || n == "drafts" ||
                   n == "gelöschte elemente" || n == "papierkorb" || n == "trash" || n == "deleted items" ||
                   n == "junk-e-mail" || n == "junk" || n == "spam" ||
                   n == "postausgang" || n == "outbox";
        }

        private static Outlook.Conversation? GetItemConversation(object item)
        {
            if (item == null) return null;
            try
            {
                if (item is Outlook._MailItem mail) return mail.GetConversation();
                if (item is Outlook._MeetingItem meeting) return meeting.GetConversation();
                if (item is Outlook._PostItem post) return post.GetConversation();

                dynamic dyn = item;
                return (Outlook.Conversation?)dyn.GetConversation();
            }
            catch
            {
                return null;
            }
        }

        private static string? GetItemConversationId(object item)
        {
            if (item == null) return null;
            try
            {
                if (item is Outlook._MailItem mail) return mail.ConversationID;
                if (item is Outlook._MeetingItem meeting) return meeting.ConversationID;
                if (item is Outlook._PostItem post) return post.ConversationID;

                dynamic dyn = item;
                return (string?)dyn.ConversationID;
            }
            catch
            {
                return null;
            }
        }

        private static string? GetItemStoreId(object item)
        {
            try
            {
                if (item is Outlook._MailItem mail)
                {
                    Outlook.MAPIFolder? parent = mail.Parent as Outlook.MAPIFolder;
                    if (parent != null)
                    {
                        string storeId = parent.StoreID;
                        Marshal.ReleaseComObject(parent);
                        return storeId;
                    }
                }
                if (item is Outlook._MeetingItem meeting)
                {
                    Outlook.MAPIFolder? parent = meeting.Parent as Outlook.MAPIFolder;
                    if (parent != null)
                    {
                        string storeId = parent.StoreID;
                        Marshal.ReleaseComObject(parent);
                        return storeId;
                    }
                }
                if (item is Outlook._PostItem post)
                {
                    Outlook.MAPIFolder? parent = post.Parent as Outlook.MAPIFolder;
                    if (parent != null)
                    {
                        string storeId = parent.StoreID;
                        Marshal.ReleaseComObject(parent);
                        return storeId;
                    }
                }

                dynamic dyn = item;
                Outlook.MAPIFolder? p = dyn.Parent as Outlook.MAPIFolder;
                if (p != null)
                {
                    string storeId = p.StoreID;
                    Marshal.ReleaseComObject(p);
                    return storeId;
                }
            }
            catch { }

            return null;
        }

        private static string? GetItemEntryId(object item)
        {
            try
            {
                if (item is Outlook._MailItem mail) return mail.EntryID;
                if (item is Outlook._MeetingItem meeting) return meeting.EntryID;
                if (item is Outlook._PostItem post) return post.EntryID;

                dynamic dyn = item;
                return (string)dyn.EntryID;
            }
            catch
            {
                return null;
            }
        }

        private static string? GetItemSubject(object item)
        {
            try
            {
                if (item is Outlook._MailItem mail) return mail.Subject;
                if (item is Outlook._MeetingItem meeting) return meeting.Subject;
                if (item is Outlook._PostItem post) return post.Subject;
                if (item is Outlook.ConversationHeader ch) return ch.ConversationTopic;

                dynamic dyn = item;
                return (string)dyn.Subject;
            }
            catch
            {
                return null;
            }
        }

        #endregion
    }
}
