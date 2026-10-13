using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using Move2Folder.Models;

namespace Move2Folder.Services
{
    public static class FolderTreeReconciler
    {
        public static void Reconcile(
            ObservableCollection<FolderItem> roots,
            List<FolderItem> flatList,
            List<FolderSnapshot> freshRoots,
            ref FolderItem? selectedFolder,
            Action<FolderItem>? onCacheItemAdded = null,
            Action<FolderItem>? onCacheItemRemoved = null,
            Action<FolderItem?>? onSelectedFolderChanged = null)
        {
            if (roots == null || freshRoots == null || freshRoots.Count == 0)
                return;

            // 1. Index all fresh snapshots across the entire tree by EntryId
            var allFreshMap = new Dictionary<string, FolderSnapshot>(StringComparer.OrdinalIgnoreCase);
            IndexFreshSnapshots(freshRoots, allFreshMap);

            // 2. Index all existing items currently in tree and flat list
            var existingMap = new Dictionary<string, FolderItem>(StringComparer.OrdinalIgnoreCase);
            foreach (var root in roots)
            {
                if (!string.IsNullOrEmpty(root.EntryId))
                    existingMap[root.EntryId] = root;
            }
            foreach (var item in flatList)
            {
                if (!string.IsNullOrEmpty(item.EntryId))
                    existingMap[item.EntryId] = item;
            }

            // 3. Reconcile root stores
            var freshRootMap = new Dictionary<string, FolderSnapshot>(StringComparer.OrdinalIgnoreCase);
            foreach (var fr in freshRoots)
            {
                string key = !string.IsNullOrEmpty(fr.EntryId) ? fr.EntryId : fr.StoreId;
                if (!string.IsNullOrEmpty(key))
                    freshRootMap[key] = fr;
            }

            // Check for removed root stores
            var currentRoots = roots.ToList();
            foreach (var root in currentRoots)
            {
                string key = !string.IsNullOrEmpty(root.EntryId) ? root.EntryId : root.StoreId;
                if (!string.IsNullOrEmpty(key) && !freshRootMap.ContainsKey(key))
                {
                    if (selectedFolder != null && IsSameOrDescendant(selectedFolder, root))
                    {
                        selectedFolder = null;
                        onSelectedFolderChanged?.Invoke(null);
                    }
                    roots.Remove(root);
                    RemoveDescendantsFromFlatList(root, flatList, existingMap, onCacheItemRemoved);
                }
            }

            // Reconcile existing roots and add any new roots
            foreach (var freshRoot in freshRoots)
            {
                var existingRoot = roots.FirstOrDefault(r =>
                    (!string.IsNullOrEmpty(freshRoot.EntryId) && string.Equals(r.EntryId, freshRoot.EntryId, StringComparison.OrdinalIgnoreCase)) ||
                    (freshRoot.IsStore && r.IsStore && string.Equals(r.StoreId, freshRoot.StoreId, StringComparison.OrdinalIgnoreCase)));

                if (existingRoot != null)
                {
                    // Update root properties if renamed
                    if (existingRoot.Name != freshRoot.Name) existingRoot.Name = freshRoot.Name;
                    if (existingRoot.FullPath != freshRoot.FullPath) existingRoot.FullPath = freshRoot.FullPath;
                    if (existingRoot.DisplayPath != freshRoot.DisplayPath) existingRoot.DisplayPath = freshRoot.DisplayPath;

                    // Reconcile children of this root
                    ReconcileChildren(
                        existingRoot,
                        freshRoot.Children,
                        allFreshMap,
                        existingMap,
                        flatList,
                        ref selectedFolder,
                        onCacheItemAdded,
                        onCacheItemRemoved,
                        onSelectedFolderChanged);
                }
                else
                {
                    // New root added
                    var newRoot = CreateFolderItemFromSnapshot(freshRoot, null, flatList, existingMap, onCacheItemAdded);
                    roots.Add(newRoot);
                }
            }
        }

        private static void ReconcileChildren(
            FolderItem parentItem,
            List<FolderSnapshot> freshChildren,
            Dictionary<string, FolderSnapshot> allFreshMap,
            Dictionary<string, FolderItem> existingMap,
            List<FolderItem> flatList,
            ref FolderItem? selectedFolder,
            Action<FolderItem>? onCacheItemAdded,
            Action<FolderItem>? onCacheItemRemoved,
            Action<FolderItem?>? onSelectedFolderChanged)
        {
            var freshChildMap = new Dictionary<string, FolderSnapshot>(StringComparer.OrdinalIgnoreCase);
            foreach (var f in freshChildren)
            {
                if (!string.IsNullOrEmpty(f.EntryId))
                    freshChildMap[f.EntryId] = f;
            }

            var currentChildren = parentItem.Children.ToList();

            // 1. Remove deleted folders (or folders that moved to another parent)
            foreach (var child in currentChildren)
            {
                if (!freshChildMap.ContainsKey(child.EntryId))
                {
                    if (!allFreshMap.ContainsKey(child.EntryId))
                    {
                        // Folder was deleted from Outlook
                        if (selectedFolder != null && IsSameOrDescendant(selectedFolder, child))
                        {
                            selectedFolder = parentItem;
                            onSelectedFolderChanged?.Invoke(parentItem);
                        }

                        parentItem.Children.Remove(child);
                        RemoveDescendantsFromFlatList(child, flatList, existingMap, onCacheItemRemoved);
                    }
                    else
                    {
                        // Folder was moved to another parent in Outlook
                        parentItem.Children.Remove(child);
                    }
                }
            }

            // 2. Update existing children that remained under this parent
            foreach (var child in parentItem.Children.ToList())
            {
                if (freshChildMap.TryGetValue(child.EntryId, out var freshSnap))
                {
                    if (child.Name != freshSnap.Name) child.Name = freshSnap.Name;
                    if (child.FullPath != freshSnap.FullPath) child.FullPath = freshSnap.FullPath;
                    if (child.DisplayPath != freshSnap.DisplayPath) child.DisplayPath = freshSnap.DisplayPath;
                    if (child.FolderType != freshSnap.FolderType) child.FolderType = freshSnap.FolderType;

                    // Recursively reconcile this child's children
                    ReconcileChildren(
                        child,
                        freshSnap.Children,
                        allFreshMap,
                        existingMap,
                        flatList,
                        ref selectedFolder,
                        onCacheItemAdded,
                        onCacheItemRemoved,
                        onSelectedFolderChanged);
                }
            }

            // 3. Integrate new folders or folders moved into this parent
            foreach (var freshSnap in freshChildren)
            {
                bool existsInParent = parentItem.Children.Any(c => string.Equals(c.EntryId, freshSnap.EntryId, StringComparison.OrdinalIgnoreCase));
                if (!existsInParent)
                {
                    if (existingMap.TryGetValue(freshSnap.EntryId, out var movedItem))
                    {
                        // Folder moved here from another parent
                        movedItem.Parent = parentItem;
                        if (movedItem.Name != freshSnap.Name) movedItem.Name = freshSnap.Name;
                        if (movedItem.FullPath != freshSnap.FullPath) movedItem.FullPath = freshSnap.FullPath;
                        if (movedItem.DisplayPath != freshSnap.DisplayPath) movedItem.DisplayPath = freshSnap.DisplayPath;
                        if (movedItem.FolderType != freshSnap.FolderType) movedItem.FolderType = freshSnap.FolderType;

                        ReconcileChildren(
                            movedItem,
                            freshSnap.Children,
                            allFreshMap,
                            existingMap,
                            flatList,
                            ref selectedFolder,
                            onCacheItemAdded,
                            onCacheItemRemoved,
                            onSelectedFolderChanged);

                        InsertAlphabetically(parentItem.Children, movedItem);
                    }
                    else
                    {
                        // Brand new folder created in Outlook
                        var newItem = CreateFolderItemFromSnapshot(freshSnap, parentItem, flatList, existingMap, onCacheItemAdded);
                        InsertAlphabetically(parentItem.Children, newItem);
                    }
                }
            }

            // 4. Ensure alphabetical sorting of children without recreating items
            SortCollectionAlphabetically(parentItem.Children);
        }

        private static void IndexFreshSnapshots(IEnumerable<FolderSnapshot> snapshots, Dictionary<string, FolderSnapshot> map)
        {
            foreach (var snap in snapshots)
            {
                if (!string.IsNullOrEmpty(snap.EntryId))
                {
                    map[snap.EntryId] = snap;
                }
                IndexFreshSnapshots(snap.Children, map);
            }
        }

        private static FolderItem CreateFolderItemFromSnapshot(
            FolderSnapshot snap,
            FolderItem? parent,
            List<FolderItem> flatList,
            Dictionary<string, FolderItem> existingMap,
            Action<FolderItem>? onCacheItemAdded)
        {
            var item = new FolderItem
            {
                Name = snap.Name,
                EntryId = snap.EntryId,
                StoreId = snap.StoreId,
                FullPath = snap.FullPath,
                DisplayPath = snap.DisplayPath,
                Parent = parent,
                IsStore = snap.IsStore,
                FolderType = snap.FolderType
            };

            if (!string.IsNullOrEmpty(item.EntryId))
            {
                existingMap[item.EntryId] = item;
            }

            if (!item.IsStore)
            {
                if (!flatList.Contains(item))
                {
                    flatList.Add(item);
                }
                onCacheItemAdded?.Invoke(item);
            }

            foreach (var childSnap in snap.Children)
            {
                var childItem = CreateFolderItemFromSnapshot(childSnap, item, flatList, existingMap, onCacheItemAdded);
                InsertAlphabetically(item.Children, childItem);
            }

            return item;
        }

        private static void RemoveDescendantsFromFlatList(
            FolderItem item,
            List<FolderItem> flatList,
            Dictionary<string, FolderItem> existingMap,
            Action<FolderItem>? onCacheItemRemoved)
        {
            if (!item.IsStore)
            {
                flatList.Remove(item);
                onCacheItemRemoved?.Invoke(item);
            }
            if (!string.IsNullOrEmpty(item.EntryId))
            {
                existingMap.Remove(item.EntryId);
            }

            foreach (var child in item.Children.ToList())
            {
                RemoveDescendantsFromFlatList(child, flatList, existingMap, onCacheItemRemoved);
            }
        }

        public static bool IsSameOrDescendant(FolderItem target, FolderItem ancestor)
        {
            FolderItem? current = target;
            while (current != null)
            {
                if (current == ancestor) return true;
                current = current.Parent;
            }
            return false;
        }

        public static void InsertAlphabetically(ObservableCollection<FolderItem> collection, FolderItem newItem)
        {
            int index = 0;
            while (index < collection.Count &&
                   string.Compare(collection[index].Name, newItem.Name, StringComparison.CurrentCultureIgnoreCase) < 0)
            {
                index++;
            }
            collection.Insert(index, newItem);
        }

        public static void SortCollectionAlphabetically(ObservableCollection<FolderItem> collection)
        {
            var sorted = collection
                .OrderBy(c => c.Name, StringComparer.CurrentCultureIgnoreCase)
                .ToList();

            for (int i = 0; i < sorted.Count; i++)
            {
                int oldIndex = collection.IndexOf(sorted[i]);
                if (oldIndex != i)
                {
                    collection.Move(oldIndex, i);
                }
            }
        }
    }
}
