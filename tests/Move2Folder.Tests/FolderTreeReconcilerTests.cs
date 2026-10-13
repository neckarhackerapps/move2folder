using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using Move2Folder.Models;
using Move2Folder.Services;
using Xunit;

namespace Move2Folder.Tests
{
    public class FolderTreeReconcilerTests
    {
        [Fact]
        public void Reconcile_RenamedFolder_UpdatesPropertiesAndMaintainsInstance()
        {
            // Arrange
            var root = new FolderItem
            {
                EntryId = "STORE1",
                StoreId = "STORE1",
                Name = "user@test.com",
                FullPath = "\\user@test.com",
                DisplayPath = "user@test.com",
                IsStore = true
            };

            var child1 = new FolderItem
            {
                EntryId = "F1",
                StoreId = "STORE1",
                Name = "OldName",
                FullPath = "\\user@test.com\\OldName",
                DisplayPath = "user@test.com / OldName",
                Parent = root,
                IsStore = false,
                IsExpanded = true
            };
            root.Children.Add(child1);

            var subChild = new FolderItem
            {
                EntryId = "F2",
                StoreId = "STORE1",
                Name = "SubProject",
                FullPath = "\\user@test.com\\OldName\\SubProject",
                DisplayPath = "user@test.com / OldName / SubProject",
                Parent = child1,
                IsStore = false
            };
            child1.Children.Add(subChild);

            var roots = new ObservableCollection<FolderItem> { root };
            var flatList = new List<FolderItem> { child1, subChild };

            var propertyChangedNames = new List<string>();
            child1.PropertyChanged += (s, e) =>
            {
                if (e.PropertyName != null)
                    propertyChangedNames.Add(e.PropertyName);
            };

            // Snapshot with F1 renamed to "NewName"
            var snapRoot = new FolderSnapshot
            {
                EntryId = "STORE1",
                StoreId = "STORE1",
                Name = "user@test.com",
                FullPath = "\\user@test.com",
                DisplayPath = "user@test.com",
                IsStore = true
            };
            var snapChild1 = new FolderSnapshot
            {
                EntryId = "F1",
                StoreId = "STORE1",
                Name = "NewName",
                FullPath = "\\user@test.com\\NewName",
                DisplayPath = "user@test.com / NewName",
                IsStore = false
            };
            var snapSubChild = new FolderSnapshot
            {
                EntryId = "F2",
                StoreId = "STORE1",
                Name = "SubProject",
                FullPath = "\\user@test.com\\NewName\\SubProject",
                DisplayPath = "user@test.com / NewName / SubProject",
                IsStore = false
            };
            snapChild1.Children.Add(snapSubChild);
            snapRoot.Children.Add(snapChild1);

            FolderItem? selectedFolder = child1;

            // Act
            FolderTreeReconciler.Reconcile(roots, flatList, new List<FolderSnapshot> { snapRoot }, ref selectedFolder);

            // Assert
            Assert.Same(child1, root.Children[0]); // Object identity preserved!
            Assert.True(child1.IsExpanded); // State preserved!
            Assert.Equal("NewName", child1.Name);
            Assert.Equal("\\user@test.com\\NewName", child1.FullPath);
            Assert.Equal("user@test.com / NewName", child1.DisplayPath);
            Assert.Contains(nameof(FolderItem.Name), propertyChangedNames);

            // Subchild path updated
            Assert.Equal("user@test.com / NewName / SubProject", subChild.DisplayPath);
        }

        [Fact]
        public void Reconcile_NewFolder_IntegratesAlphabeticallyIntoTreeAndFlatList()
        {
            // Arrange
            var root = new FolderItem
            {
                EntryId = "STORE1",
                StoreId = "STORE1",
                Name = "Mailbox",
                FullPath = "\\Mailbox",
                DisplayPath = "Mailbox",
                IsStore = true
            };

            var fA = new FolderItem
            {
                EntryId = "A",
                StoreId = "STORE1",
                Name = "Apple",
                FullPath = "\\Mailbox\\Apple",
                DisplayPath = "Mailbox / Apple",
                Parent = root
            };
            var fZ = new FolderItem
            {
                EntryId = "Z",
                StoreId = "STORE1",
                Name = "Zebra",
                FullPath = "\\Mailbox\\Zebra",
                DisplayPath = "Mailbox / Zebra",
                Parent = root
            };
            root.Children.Add(fA);
            root.Children.Add(fZ);

            var roots = new ObservableCollection<FolderItem> { root };
            var flatList = new List<FolderItem> { fA, fZ };

            // Snapshot has a new folder "Mango" between Apple and Zebra
            var snapRoot = new FolderSnapshot
            {
                EntryId = "STORE1",
                StoreId = "STORE1",
                Name = "Mailbox",
                FullPath = "\\Mailbox",
                DisplayPath = "Mailbox",
                IsStore = true
            };
            var snapA = new FolderSnapshot { EntryId = "A", Name = "Apple", FullPath = "\\Mailbox\\Apple", DisplayPath = "Mailbox / Apple" };
            var snapM = new FolderSnapshot { EntryId = "M", Name = "Mango", FullPath = "\\Mailbox\\Mango", DisplayPath = "Mailbox / Mango" };
            var snapZ = new FolderSnapshot { EntryId = "Z", Name = "Zebra", FullPath = "\\Mailbox\\Zebra", DisplayPath = "Mailbox / Zebra" };

            snapRoot.Children.Add(snapA);
            snapRoot.Children.Add(snapM);
            snapRoot.Children.Add(snapZ);

            FolderItem? selectedFolder = null;
            var addedItems = new List<FolderItem>();

            // Act
            FolderTreeReconciler.Reconcile(
                roots,
                flatList,
                new List<FolderSnapshot> { snapRoot },
                ref selectedFolder,
                onCacheItemAdded: item => addedItems.Add(item));

            // Assert
            Assert.Equal(3, root.Children.Count);
            Assert.Equal("Apple", root.Children[0].Name);
            Assert.Equal("Mango", root.Children[1].Name);
            Assert.Equal("Zebra", root.Children[2].Name);

            // Added to flat list and cache
            Assert.Contains(flatList, f => f.EntryId == "M" && f.Name == "Mango");
            Assert.Contains(addedItems, f => f.EntryId == "M");
        }

        [Fact]
        public void Reconcile_DeletedFolder_RemovesFromTreeAndFlatListAndFallsBackSelection()
        {
            // Arrange
            var root = new FolderItem
            {
                EntryId = "STORE1",
                StoreId = "STORE1",
                Name = "Mailbox",
                FullPath = "\\Mailbox",
                DisplayPath = "Mailbox",
                IsStore = true
            };

            var fKeep = new FolderItem
            {
                EntryId = "F_KEEP",
                StoreId = "STORE1",
                Name = "Inbox",
                FullPath = "\\Mailbox\\Inbox",
                DisplayPath = "Mailbox / Inbox",
                Parent = root
            };
            var fDelete = new FolderItem
            {
                EntryId = "F_DEL",
                StoreId = "STORE1",
                Name = "Temporary",
                FullPath = "\\Mailbox\\Temporary",
                DisplayPath = "Mailbox / Temporary",
                Parent = root
            };
            var fDeleteSub = new FolderItem
            {
                EntryId = "F_DEL_SUB",
                StoreId = "STORE1",
                Name = "TempSub",
                FullPath = "\\Mailbox\\Temporary\\TempSub",
                DisplayPath = "Mailbox / Temporary / TempSub",
                Parent = fDelete
            };
            fDelete.Children.Add(fDeleteSub);

            root.Children.Add(fKeep);
            root.Children.Add(fDelete);

            var roots = new ObservableCollection<FolderItem> { root };
            var flatList = new List<FolderItem> { fKeep, fDelete, fDeleteSub };

            // User had the subfolder selected
            FolderItem? selectedFolder = fDeleteSub;
            FolderItem? fallbackSelection = null;
            var removedItems = new List<FolderItem>();

            // Snapshot only contains Inbox
            var snapRoot = new FolderSnapshot
            {
                EntryId = "STORE1",
                StoreId = "STORE1",
                Name = "Mailbox",
                FullPath = "\\Mailbox",
                DisplayPath = "Mailbox",
                IsStore = true
            };
            snapRoot.Children.Add(new FolderSnapshot { EntryId = "F_KEEP", Name = "Inbox", FullPath = "\\Mailbox\\Inbox", DisplayPath = "Mailbox / Inbox" });

            // Act
            FolderTreeReconciler.Reconcile(
                roots,
                flatList,
                new List<FolderSnapshot> { snapRoot },
                ref selectedFolder,
                onCacheItemRemoved: item => removedItems.Add(item),
                onSelectedFolderChanged: item => fallbackSelection = item);

            // Assert
            Assert.Single(root.Children);
            Assert.Equal("Inbox", root.Children[0].Name);

            // Both deleted folder and its child removed from flat list
            Assert.Single(flatList);
            Assert.Equal("Inbox", flatList[0].Name);

            // Both reported to cache cleanup
            Assert.Contains(removedItems, f => f.EntryId == "F_DEL");
            Assert.Contains(removedItems, f => f.EntryId == "F_DEL_SUB");

            // Selection fell back to parent root
            Assert.Same(root, selectedFolder);
            Assert.Same(root, fallbackSelection);
        }

        [Fact]
        public void Reconcile_MovedFolder_MovesToNewParentAndPreservesSubtree()
        {
            // Arrange
            var root = new FolderItem
            {
                EntryId = "STORE1",
                StoreId = "STORE1",
                Name = "Mailbox",
                FullPath = "\\Mailbox",
                DisplayPath = "Mailbox",
                IsStore = true
            };

            var fParentA = new FolderItem
            {
                EntryId = "P_A",
                StoreId = "STORE1",
                Name = "FolderA",
                FullPath = "\\Mailbox\\FolderA",
                DisplayPath = "Mailbox / FolderA",
                Parent = root
            };
            var fParentB = new FolderItem
            {
                EntryId = "P_B",
                StoreId = "STORE1",
                Name = "FolderB",
                FullPath = "\\Mailbox\\FolderB",
                DisplayPath = "Mailbox / FolderB",
                Parent = root
            };

            var fMove = new FolderItem
            {
                EntryId = "F_MOVE",
                StoreId = "STORE1",
                Name = "MovedChild",
                FullPath = "\\Mailbox\\FolderA\\MovedChild",
                DisplayPath = "Mailbox / FolderA / MovedChild",
                Parent = fParentA,
                IsExpanded = true
            };
            fParentA.Children.Add(fMove);

            root.Children.Add(fParentA);
            root.Children.Add(fParentB);

            var roots = new ObservableCollection<FolderItem> { root };
            var flatList = new List<FolderItem> { fParentA, fParentB, fMove };

            // Snapshot has MovedChild under FolderB instead of FolderA
            var snapRoot = new FolderSnapshot
            {
                EntryId = "STORE1",
                StoreId = "STORE1",
                Name = "Mailbox",
                FullPath = "\\Mailbox",
                DisplayPath = "Mailbox",
                IsStore = true
            };
            var snapParentA = new FolderSnapshot { EntryId = "P_A", Name = "FolderA", FullPath = "\\Mailbox\\FolderA", DisplayPath = "Mailbox / FolderA" };
            var snapParentB = new FolderSnapshot { EntryId = "P_B", Name = "FolderB", FullPath = "\\Mailbox\\FolderB", DisplayPath = "Mailbox / FolderB" };
            var snapMove = new FolderSnapshot { EntryId = "F_MOVE", Name = "MovedChild", FullPath = "\\Mailbox\\FolderB\\MovedChild", DisplayPath = "Mailbox / FolderB / MovedChild" };

            snapParentB.Children.Add(snapMove);
            snapRoot.Children.Add(snapParentA);
            snapRoot.Children.Add(snapParentB);

            FolderItem? selectedFolder = fMove;

            // Act
            FolderTreeReconciler.Reconcile(roots, flatList, new List<FolderSnapshot> { snapRoot }, ref selectedFolder);

            // Assert
            Assert.Empty(fParentA.Children);
            Assert.Single(fParentB.Children);
            Assert.Same(fMove, fParentB.Children[0]); // Same instance!
            Assert.Same(fParentB, fMove.Parent);
            Assert.True(fMove.IsExpanded);
            Assert.Equal("Mailbox / FolderB / MovedChild", fMove.DisplayPath);
        }

        [Fact]
        public void Reconcile_RenamedFolder_ReordersAlphabeticallyWithoutRecreatingNodes()
        {
            // Arrange: Apple, Zebra
            var root = new FolderItem { EntryId = "S", Name = "Store", IsStore = true };
            var fApple = new FolderItem { EntryId = "1", Name = "Apple", Parent = root };
            var fZebra = new FolderItem { EntryId = "2", Name = "Zebra", Parent = root };
            root.Children.Add(fApple);
            root.Children.Add(fZebra);

            var roots = new ObservableCollection<FolderItem> { root };
            var flatList = new List<FolderItem> { fApple, fZebra };

            // Rename "Zebra" to "Aardvark" -> should move to index 0 before Apple
            var snapRoot = new FolderSnapshot { EntryId = "S", Name = "Store", IsStore = true };
            snapRoot.Children.Add(new FolderSnapshot { EntryId = "1", Name = "Apple" });
            snapRoot.Children.Add(new FolderSnapshot { EntryId = "2", Name = "Aardvark" });

            FolderItem? selectedFolder = fZebra;

            // Act
            FolderTreeReconciler.Reconcile(roots, flatList, new List<FolderSnapshot> { snapRoot }, ref selectedFolder);

            // Assert
            Assert.Equal(2, root.Children.Count);
            Assert.Same(fZebra, root.Children[0]); // Object identity preserved!
            Assert.Equal("Aardvark", root.Children[0].Name);
            Assert.Same(fApple, root.Children[1]);
            Assert.Equal("Apple", root.Children[1].Name);
        }

        [Fact]
        public void Reconcile_MultipleChangesSimultaneously_CorrectlyUpdatesTree()
        {
            // Root has Folder1, Folder2, Folder3
            var root = new FolderItem { EntryId = "S", Name = "Store", IsStore = true };
            var f1 = new FolderItem { EntryId = "1", Name = "Folder1", Parent = root };
            var f2 = new FolderItem { EntryId = "2", Name = "Folder2", Parent = root };
            var f3 = new FolderItem { EntryId = "3", Name = "Folder3", Parent = root };
            root.Children.Add(f1);
            root.Children.Add(f2);
            root.Children.Add(f3);

            var roots = new ObservableCollection<FolderItem> { root };
            var flatList = new List<FolderItem> { f1, f2, f3 };

            // In snapshot:
            // - Folder1 is deleted
            // - Folder2 is renamed to "Beta"
            // - Folder3 is unchanged
            // - Folder4 "Alpha" is new
            var snapRoot = new FolderSnapshot { EntryId = "S", Name = "Store", IsStore = true };
            snapRoot.Children.Add(new FolderSnapshot { EntryId = "4", Name = "Alpha" });
            snapRoot.Children.Add(new FolderSnapshot { EntryId = "2", Name = "Beta" });
            snapRoot.Children.Add(new FolderSnapshot { EntryId = "3", Name = "Folder3" });

            FolderItem? selectedFolder = f1; // Was on deleted folder
            FolderItem? selectionChanged = null;

            // Act
            FolderTreeReconciler.Reconcile(
                roots,
                flatList,
                new List<FolderSnapshot> { snapRoot },
                ref selectedFolder,
                onSelectedFolderChanged: item => selectionChanged = item);

            // Assert
            Assert.Equal(3, root.Children.Count);
            Assert.Equal("Alpha", root.Children[0].Name);
            Assert.Equal("Beta", root.Children[1].Name);
            Assert.Equal("Folder3", root.Children[2].Name);

            // f2 was preserved
            Assert.Same(f2, root.Children[1]);
            Assert.Same(f3, root.Children[2]);

            // f1 removed from flatList
            Assert.DoesNotContain(flatList, f => f.EntryId == "1");
            Assert.Contains(flatList, f => f.EntryId == "4");

            // Selected folder fell back to parent root
            Assert.Same(root, selectedFolder);
            Assert.Same(root, selectionChanged);
        }
    }
}
