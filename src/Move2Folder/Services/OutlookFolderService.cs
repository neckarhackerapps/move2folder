using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Runtime.InteropServices;
using System.Threading;
using Move2Folder.Models;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace Move2Folder.Services
{
    public class OutlookFolderService
    {
        private readonly Outlook.Application _outlookApp;
        public Outlook.Application OutlookApp => _outlookApp;
        private readonly Dictionary<string, (ObservableCollection<FolderItem> Roots, List<FolderItem> FlatList)> _cache =
            new Dictionary<string, (ObservableCollection<FolderItem> Roots, List<FolderItem> FlatList)>(StringComparer.OrdinalIgnoreCase);

        public OutlookFolderService(Outlook.Application outlookApp)
        {
            _outlookApp = outlookApp ?? throw new ArgumentNullException(nameof(outlookApp));
        }

        public (ObservableCollection<FolderItem> Roots, List<FolderItem> FlatList) LoadFolders(bool forceRefresh = false)
        {
            return LoadFolders(targetStoreId: null, forceRefresh: forceRefresh);
        }

        public (ObservableCollection<FolderItem> Roots, List<FolderItem> FlatList) LoadFolders(string? targetStoreId, bool forceRefresh = false)
        {
            string cacheKey = targetStoreId ?? "ALL";

            if (!forceRefresh && _cache.TryGetValue(cacheKey, out var cachedData))
            {
                return cachedData;
            }

            var roots = new ObservableCollection<FolderItem>();
            var flatList = new List<FolderItem>();

            Outlook.NameSpace? session = null;
            Outlook.Folders? sessionFolders = null;

            try
            {
                session = _outlookApp.Session;
                sessionFolders = session.Folders;

                int rootCount = sessionFolders.Count;

                // Pass 1: If targetStoreId is specified, search for matching store root among session.Folders
                if (!string.IsNullOrEmpty(targetStoreId))
                {
                    for (int i = 1; i <= rootCount; i++)
                    {
                        Outlook.MAPIFolder? rootFolder = null;
                        try
                        {
                            rootFolder = sessionFolders[i];
                            if (string.Equals(rootFolder.StoreID, targetStoreId, StringComparison.OrdinalIgnoreCase))
                            {
                                var rootItem = BuildFolderHierarchy(rootFolder, null, flatList, isStoreRoot: true);
                                if (rootItem != null)
                                {
                                    roots.Add(rootItem);
                                }
                                break;
                            }
                        }
                        catch
                        {
                            // Ignore inaccessible stores
                        }
                        finally
                        {
                            if (rootFolder != null)
                                Marshal.ReleaseComObject(rootFolder);
                        }
                    }

                    // If not found in session.Folders, search in session.Stores
                    if (roots.Count == 0)
                    {
                        Outlook.Stores? stores = null;
                        try
                        {
                            stores = session.Stores;
                            int storeCount = stores.Count;
                            for (int s = 1; s <= storeCount; s++)
                            {
                                Outlook.Store? store = null;
                                Outlook.MAPIFolder? rootFolder = null;
                                try
                                {
                                    store = stores[s];
                                    if (string.Equals(store.StoreID, targetStoreId, StringComparison.OrdinalIgnoreCase))
                                    {
                                        rootFolder = store.GetRootFolder();
                                        if (rootFolder != null)
                                        {
                                            var rootItem = BuildFolderHierarchy(rootFolder, null, flatList, isStoreRoot: true);
                                            if (rootItem != null)
                                            {
                                                roots.Add(rootItem);
                                            }
                                            break;
                                        }
                                    }
                                }
                                catch { }
                                finally
                                {
                                    if (rootFolder != null) Marshal.ReleaseComObject(rootFolder);
                                    if (store != null) Marshal.ReleaseComObject(store);
                                }
                            }
                        }
                        catch { }
                        finally
                        {
                            if (stores != null) Marshal.ReleaseComObject(stores);
                        }
                    }
                }

                // Pass 2: Fallback if targetStoreId was not found or not specified: load all stores
                if (roots.Count == 0)
                {
                    for (int i = 1; i <= rootCount; i++)
                    {
                        Outlook.MAPIFolder? rootFolder = null;
                        try
                        {
                            rootFolder = sessionFolders[i];
                            var rootItem = BuildFolderHierarchy(rootFolder, null, flatList, isStoreRoot: true);
                            if (rootItem != null)
                            {
                                roots.Add(rootItem);
                            }
                        }
                        catch
                        {
                            // Ignore inaccessible stores (e.g. offline shared mailboxes)
                        }
                        finally
                        {
                            if (rootFolder != null)
                                Marshal.ReleaseComObject(rootFolder);
                        }
                    }
                }
            }
            finally
            {
                if (sessionFolders != null)
                    Marshal.ReleaseComObject(sessionFolders);
                if (session != null)
                    Marshal.ReleaseComObject(session);
            }

            var result = (roots, flatList);
            _cache[cacheKey] = result;

            return result;
        }

        private FolderItem? BuildFolderHierarchy(
            Outlook.MAPIFolder folder,
            FolderItem? parent,
            List<FolderItem> flatList,
            bool isStoreRoot = false)
        {
            try
            {
                Outlook.OlItemType itemType = folder.DefaultItemType;

                // Only include mail folders or folders that have subfolders
                Outlook.Folders subFolders = folder.Folders;
                int subFolderCount = subFolders.Count;

                bool isMailFolder = itemType == Outlook.OlItemType.olMailItem;
                bool hasSubFolders = subFolderCount > 0;

                // Skip non-mail leaf folders (e.g. Contacts, Calendar, Notes, Tasks)
                if (!isStoreRoot && !isMailFolder && !hasSubFolders)
                {
                    Marshal.ReleaseComObject(subFolders);
                    return null;
                }

                string name = folder.Name ?? LocalizationService.UntitledFolder;
                string entryId = folder.EntryID ?? string.Empty;
                string storeId = folder.StoreID ?? string.Empty;
                string fullPath = folder.FolderPath ?? name;

                string displayPath = fullPath.TrimStart('\\').Replace("\\", " / ");

                var item = new FolderItem
                {
                    Name = name,
                    EntryId = entryId,
                    StoreId = storeId,
                    FullPath = fullPath,
                    DisplayPath = displayPath,
                    Parent = parent,
                    IsStore = isStoreRoot,
                    FolderType = DetermineFolderType(name, isStoreRoot, isMailFolder)
                };

                // Store roots are headers; subfolders are valid move destinations
                if (!isStoreRoot)
                {
                    flatList.Add(item);
                }

                if (subFolderCount > 0)
                {
                    for (int i = 1; i <= subFolderCount; i++)
                    {
                        Outlook.MAPIFolder? childFolder = null;
                        try
                        {
                            childFolder = subFolders[i];
                            var childItem = BuildFolderHierarchy(childFolder, item, flatList, isStoreRoot: false);
                            if (childItem != null)
                            {
                                item.Children.Add(childItem);
                            }
                        }
                        catch
                        {
                            // Skip inaccessible child folders
                        }
                        finally
                        {
                            if (childFolder != null)
                                Marshal.ReleaseComObject(childFolder);
                        }
                    }
                }

                Marshal.ReleaseComObject(subFolders);
                return item;
            }
            catch
            {
                return null;
            }
        }

        public static string DetermineFolderType(string name, bool isStoreRoot, bool isMailFolder)
        {
            if (isStoreRoot)
                return "Store";

            string lower = name.ToLowerInvariant();
            if (lower.Contains("inbox") || lower.Contains("posteingang"))
                return "Inbox";
            if (lower.Contains("sent") || lower.Contains("gesendet"))
                return "Sent";
            if (lower.Contains("archive") || lower.Contains("archiv"))
                return "Archive";
            if (lower.Contains("trash") || lower.Contains("papierkorb") || lower.Contains("gelöscht"))
                return "Trash";
            if (lower.Contains("draft") || lower.Contains("entwurf") || lower.Contains("entwürfe"))
                return "Drafts";
            if (lower.Contains("junk") || lower.Contains("spam"))
                return "Junk";

            return isMailFolder ? "Folder" : "Special";
        }

        public Outlook.MAPIFolder? GetFolder(string entryId, string storeId)
        {
            if (string.IsNullOrEmpty(entryId))
                return null;

            Outlook.NameSpace? session = null;
            try
            {
                session = _outlookApp.Session;
                if (!string.IsNullOrEmpty(storeId))
                {
                    return session.GetFolderFromID(entryId, storeId);
                }
                return session.GetFolderFromID(entryId);
            }
            catch
            {
                return null;
            }
            finally
            {
                if (session != null)
                    Marshal.ReleaseComObject(session);
            }
        }

        public int MoveItems(IEnumerable<object> items, Outlook.MAPIFolder targetFolder)
        {
            int movedCount = 0;
            foreach (var item in items)
            {
                try
                {
                    if (item is Outlook._MailItem mail)
                    {
                        mail.Move(targetFolder);
                        movedCount++;
                    }
                    else if (item is Outlook._MeetingItem meeting)
                    {
                        meeting.Move(targetFolder);
                        movedCount++;
                    }
                    else if (item is Outlook._PostItem post)
                    {
                        post.Move(targetFolder);
                        movedCount++;
                    }
                    else
                    {
                        // Dynamic dispatch fallback for other Outlook items supporting Move
                        dynamic dynamicItem = item;
                        dynamicItem.Move(targetFolder);
                        movedCount++;
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Error moving item: {ex.Message}");
                }
                finally
                {
                    try
                    {
                        Marshal.ReleaseComObject(item);
                    }
                    catch { }
                }
            }

            return movedCount;
        }

        public FolderItem CreateFolder(FolderItem parentItem, string folderName)
        {
            if (parentItem == null)
                throw new ArgumentNullException(nameof(parentItem));
            if (string.IsNullOrWhiteSpace(folderName))
                throw new ArgumentException("Folder name cannot be empty", nameof(folderName));

            string trimmedName = folderName.Trim();
            Outlook.MAPIFolder? parentFolder = null;
            Outlook.Folders? subFolders = null;
            Outlook.MAPIFolder? newOutlookFolder = null;

            try
            {
                parentFolder = GetFolder(parentItem.EntryId, parentItem.StoreId);
                if (parentFolder == null && parentItem.IsStore)
                {
                    parentFolder = GetStoreRootFolder(parentItem.StoreId);
                }

                if (parentFolder == null)
                {
                    throw new InvalidOperationException(
                        LocalizationService.MsgFolderOpenError(parentItem.Name));
                }

                subFolders = parentFolder.Folders;

                try
                {
                    newOutlookFolder = subFolders.Add(trimmedName, Outlook.OlDefaultFolders.olFolderInbox);
                }
                catch
                {
                    newOutlookFolder = subFolders.Add(trimmedName, Type.Missing);
                }

                string entryId = newOutlookFolder.EntryID ?? string.Empty;
                string storeId = newOutlookFolder.StoreID ?? parentItem.StoreId;
                string fullPath = newOutlookFolder.FolderPath ?? (parentItem.FullPath + "\\" + trimmedName);
                string displayPath = fullPath.TrimStart('\\').Replace("\\", " / ");

                var newItem = new FolderItem
                {
                    Name = newOutlookFolder.Name ?? trimmedName,
                    EntryId = entryId,
                    StoreId = storeId,
                    FullPath = fullPath,
                    DisplayPath = displayPath,
                    Parent = parentItem,
                    IsStore = false,
                    FolderType = DetermineFolderType(newOutlookFolder.Name ?? trimmedName, isStoreRoot: false, isMailFolder: true)
                };

                // Insert in alphabetical order among children
                int insertIndex = 0;
                while (insertIndex < parentItem.Children.Count &&
                       string.Compare(parentItem.Children[insertIndex].Name, newItem.Name, StringComparison.CurrentCultureIgnoreCase) < 0)
                {
                    insertIndex++;
                }
                parentItem.Children.Insert(insertIndex, newItem);

                // Also add to cached FlatList if available
                string cacheKey = parentItem.StoreId;
                if (!string.IsNullOrEmpty(cacheKey) && _cache.TryGetValue(cacheKey, out var cachedData))
                {
                    if (!cachedData.FlatList.Contains(newItem))
                    {
                        cachedData.FlatList.Add(newItem);
                    }
                }
                if (_cache.TryGetValue("ALL", out var allCachedData))
                {
                    if (!allCachedData.FlatList.Contains(newItem))
                    {
                        allCachedData.FlatList.Add(newItem);
                    }
                }

                return newItem;
            }
            finally
            {
                if (newOutlookFolder != null)
                    Marshal.ReleaseComObject(newOutlookFolder);
                if (subFolders != null)
                    Marshal.ReleaseComObject(subFolders);
                if (parentFolder != null)
                    Marshal.ReleaseComObject(parentFolder);
            }
        }

        private Outlook.MAPIFolder? GetStoreRootFolder(string storeId)
        {
            Outlook.NameSpace? session = null;
            Outlook.Stores? stores = null;
            try
            {
                session = _outlookApp.Session;
                stores = session.Stores;
                int count = stores.Count;
                for (int i = 1; i <= count; i++)
                {
                    Outlook.Store? store = null;
                    try
                    {
                        store = stores[i];
                        if (string.Equals(store.StoreID, storeId, StringComparison.OrdinalIgnoreCase))
                        {
                            return store.GetRootFolder();
                        }
                    }
                    finally
                    {
                        if (store != null)
                            Marshal.ReleaseComObject(store);
                    }
                }
            }
            catch { }
            finally
            {
                if (stores != null)
                    Marshal.ReleaseComObject(stores);
                if (session != null)
                    Marshal.ReleaseComObject(session);
            }
            return null;
        }

        public static List<FolderSnapshot> FetchFolderSnapshots(
            Outlook.Application outlookApp,
            string? targetStoreId,
            CancellationToken cancellationToken)
        {
            var roots = new List<FolderSnapshot>();
            Outlook.NameSpace? session = null;
            Outlook.Folders? sessionFolders = null;

            try
            {
                if (cancellationToken.IsCancellationRequested)
                    return roots;

                session = outlookApp.Session;
                sessionFolders = session.Folders;
                int rootCount = sessionFolders.Count;

                // Pass 1: If targetStoreId is specified, search for matching store root among session.Folders
                if (!string.IsNullOrEmpty(targetStoreId))
                {
                    for (int i = 1; i <= rootCount; i++)
                    {
                        if (cancellationToken.IsCancellationRequested)
                            return roots;

                        Outlook.MAPIFolder? rootFolder = null;
                        try
                        {
                            rootFolder = sessionFolders[i];
                            if (string.Equals(rootFolder.StoreID, targetStoreId, StringComparison.OrdinalIgnoreCase))
                            {
                                var rootSnap = BuildFolderSnapshot(rootFolder, cancellationToken, isStoreRoot: true);
                                if (rootSnap != null)
                                {
                                    roots.Add(rootSnap);
                                }
                                break;
                            }
                        }
                        catch { }
                        finally
                        {
                            if (rootFolder != null)
                                Marshal.ReleaseComObject(rootFolder);
                        }
                    }

                    // If not found in session.Folders, search in session.Stores
                    if (roots.Count == 0 && !cancellationToken.IsCancellationRequested)
                    {
                        Outlook.Stores? stores = null;
                        try
                        {
                            stores = session.Stores;
                            int storeCount = stores.Count;
                            for (int s = 1; s <= storeCount; s++)
                            {
                                if (cancellationToken.IsCancellationRequested)
                                    return roots;

                                Outlook.Store? store = null;
                                Outlook.MAPIFolder? rootFolder = null;
                                try
                                {
                                    store = stores[s];
                                    if (string.Equals(store.StoreID, targetStoreId, StringComparison.OrdinalIgnoreCase))
                                    {
                                        rootFolder = store.GetRootFolder();
                                        if (rootFolder != null)
                                        {
                                            var rootSnap = BuildFolderSnapshot(rootFolder, cancellationToken, isStoreRoot: true);
                                            if (rootSnap != null)
                                            {
                                                roots.Add(rootSnap);
                                            }
                                            break;
                                        }
                                    }
                                }
                                catch { }
                                finally
                                {
                                    if (rootFolder != null) Marshal.ReleaseComObject(rootFolder);
                                    if (store != null) Marshal.ReleaseComObject(store);
                                }
                            }
                        }
                        catch { }
                        finally
                        {
                            if (stores != null) Marshal.ReleaseComObject(stores);
                        }
                    }
                }

                // Pass 2: Fallback if targetStoreId was not found or not specified: load all stores
                if (roots.Count == 0 && !cancellationToken.IsCancellationRequested)
                {
                    for (int i = 1; i <= rootCount; i++)
                    {
                        if (cancellationToken.IsCancellationRequested)
                            return roots;

                        Outlook.MAPIFolder? rootFolder = null;
                        try
                        {
                            rootFolder = sessionFolders[i];
                            var rootSnap = BuildFolderSnapshot(rootFolder, cancellationToken, isStoreRoot: true);
                            if (rootSnap != null)
                            {
                                roots.Add(rootSnap);
                            }
                        }
                        catch { }
                        finally
                        {
                            if (rootFolder != null)
                                Marshal.ReleaseComObject(rootFolder);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error fetching folder snapshots: {ex.Message}");
            }
            finally
            {
                if (sessionFolders != null)
                    Marshal.ReleaseComObject(sessionFolders);
                if (session != null)
                    Marshal.ReleaseComObject(session);
            }

            return roots;
        }

        private static FolderSnapshot? BuildFolderSnapshot(
            Outlook.MAPIFolder folder,
            CancellationToken cancellationToken,
            bool isStoreRoot = false)
        {
            if (cancellationToken.IsCancellationRequested)
                return null;

            try
            {
                Outlook.OlItemType itemType = folder.DefaultItemType;

                // Only include mail folders or folders that have subfolders
                Outlook.Folders subFolders = folder.Folders;
                int subFolderCount = subFolders.Count;

                bool isMailFolder = itemType == Outlook.OlItemType.olMailItem;
                bool hasSubFolders = subFolderCount > 0;

                // Skip non-mail leaf folders (e.g. Contacts, Calendar, Notes, Tasks)
                if (!isStoreRoot && !isMailFolder && !hasSubFolders)
                {
                    Marshal.ReleaseComObject(subFolders);
                    return null;
                }

                string name = folder.Name ?? LocalizationService.UntitledFolder;
                string entryId = folder.EntryID ?? string.Empty;
                string storeId = folder.StoreID ?? string.Empty;
                string fullPath = folder.FolderPath ?? name;
                string displayPath = fullPath.TrimStart('\\').Replace("\\", " / ");

                var snapshot = new FolderSnapshot
                {
                    Name = name,
                    EntryId = entryId,
                    StoreId = storeId,
                    FullPath = fullPath,
                    DisplayPath = displayPath,
                    IsStore = isStoreRoot,
                    FolderType = DetermineFolderType(name, isStoreRoot, isMailFolder)
                };

                if (subFolderCount > 0)
                {
                    for (int i = 1; i <= subFolderCount; i++)
                    {
                        if (cancellationToken.IsCancellationRequested)
                            break;

                        Outlook.MAPIFolder? childFolder = null;
                        try
                        {
                            childFolder = subFolders[i];
                            var childSnap = BuildFolderSnapshot(childFolder, cancellationToken, isStoreRoot: false);
                            if (childSnap != null)
                            {
                                snapshot.Children.Add(childSnap);
                            }
                        }
                        catch { }
                        finally
                        {
                            if (childFolder != null)
                                Marshal.ReleaseComObject(childFolder);
                        }
                    }
                }

                Marshal.ReleaseComObject(subFolders);
                return snapshot;
            }
            catch
            {
                return null;
            }
        }

        public void UpdateCacheOnItemAdded(FolderItem item)
        {
            foreach (var entry in _cache.Values)
            {
                if (!entry.FlatList.Contains(item))
                {
                    entry.FlatList.Add(item);
                }
            }
        }

        public void UpdateCacheOnItemRemoved(FolderItem item)
        {
            foreach (var entry in _cache.Values)
            {
                entry.FlatList.Remove(item);
            }
        }
    }
}
