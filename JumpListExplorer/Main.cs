using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using JumpListExplorer.Shell;
using JumpListExplorer.Utilities;

namespace JumpListExplorer
{
    public partial class Main : Form
    {
        public Main()
        {
            InitializeComponent();
            Icon = Resources.JumpListExplorerIcon;
            listViewMain.ListViewItemSorter = new ColumnSorter { SortColumn = columnHeaderDate.Index, Order = SortOrder.Descending };
            LoadApplicationUserModelIDs();
            listViewMain.Sort();
            listViewMain.ColumnClick += (s, e) => { ((ColumnSorter)listViewMain.ListViewItemSorter).HandleClick(listViewMain, e.Column); };
        }

        private void LoadApplicationUserModelIDs()
        {
            listViewMain.Items.Clear();
            foreach (var aumid in AutomaticDestinationList.EnumerateAppUserModelIDs().OrderBy(i => i))
            {
                foreach (var item in AutomaticDestinationList.GetItems(aumid).OrderBy(i => i.SIGDN_DESKTOPABSOLUTEPARSING))
                {
                    var imageIndex = -1;
                    try
                    {
                        var path = item.SIGDN_FILESYSPATH;
                        if (!string.IsNullOrWhiteSpace(path))
                        {
                            var extension = Path.GetExtension(path).ToLowerInvariant();
                            if (!string.IsNullOrEmpty(extension))
                            {
                                imageIndex = imageListIcons.Images.IndexOfKey(extension);
                                if (imageIndex < 0)
                                {
                                    var icon = item.GetIconFromImageList(Interop.SHIL.SHIL_SYSSMALL);
                                    if (icon != null)
                                    {
                                        imageIndex = imageListIcons.Images.Count;
                                        imageListIcons.Images.Add(extension, icon);
                                    }
                                }
                            }
                        }
                    }
                    catch
                    {
                        // continue
                    }

                    ListViewItem lvi;
                    if (imageIndex >= 0)
                    {
                        lvi = listViewMain.Items.Add(aumid, imageIndex);
                    }
                    else
                    {
                        lvi = listViewMain.Items.Add(aumid);
                    }

                    lvi.Tag = new Model(aumid, item);
                    lvi.SubItems.Add(item.DateAccessed?.ToString("yyyy/MM/dd HH:mm:ss"));
                    lvi.SubItems.Add(item.SIGDN_NORMALDISPLAY);
                    lvi.SubItems.Add(item.SIGDN_DESKTOPABSOLUTEPARSING);
                }
            }

            listViewMain.AutoResizeColumns(ColumnHeaderAutoResizeStyle.ColumnContent);
            columnHeaderName.Width = 120;
            Text = $"{WinformsUtilities.ApplicationName} - {listViewMain.Items.Count} items";
        }

        private void RemoveAllDeleted()
        {
            foreach (var aumid in AutomaticDestinationList.EnumerateAppUserModelIDs().OrderBy(i => i))
            {
                var list = new List<Item>();
                foreach (var item in AutomaticDestinationList.GetItems(aumid).OrderBy(i => i.SIGDN_DESKTOPABSOLUTEPARSING))
                {
                    var path = item.SIGDN_FILESYSPATH;
                    if (path != null && !Directory.Exists(item.SIGDN_FILESYSPATH) && !File.Exists(item.SIGDN_FILESYSPATH))
                    {
                        list.Add(item);
                    }
                }

                AutomaticDestinationList.RemoveItems(aumid, list);
            }

            LoadApplicationUserModelIDs();
        }

        private List<Model> GetSelection()
        {
            var list = new List<Model>();
            foreach (var lvi in listViewMain.SelectedItems.OfType<ListViewItem>())
            {
                if (lvi.Tag is Model model)
                {
                    list.Add(model);
                }
            }
            return list;
        }

        private string? GetParentPath()
        {
            var list = GetSelection();
            if (list.Count == 1 && list[0].Item.Parent?.SIGDN_FILESYSPATH != null)
                return list[0].Item.Parent?.SIGDN_FILESYSPATH;

            return null;
        }

        private string? GetPath()
        {
            var list = GetSelection();
            if (list.Count == 1 && list[0].Item.SIGDN_FILESYSPATH != null)
                return list[0].Item.SIGDN_FILESYSPATH;

            return null;
        }

        private void ExitToolStripMenuItem_Click(object sender, EventArgs e) => Close();
        private void AboutWindowsJumpListExplorerToolStripMenuItem_Click(object sender, EventArgs e) => this.ShowMessage(Assembly.GetEntryAssembly()!.GetCustomAttribute<AssemblyTitleAttribute>()!.Title + " - " + (IntPtr.Size == 4 ? "32" : "64") + "-bit" + Environment.NewLine + "Copyright (C) 2022-" + DateTime.Now.Year + " Simon Mourier. All rights reserved.");
        private void RefreshToolStripMenuItem_Click(object sender, EventArgs e) => LoadApplicationUserModelIDs();
        private void OpenLocationToolStripMenuItem_Click(object sender, EventArgs e) => WindowsUtilities.OpenExplorer(GetParentPath());
        private void ListViewMain_DoubleClick(object sender, EventArgs e) => WindowsUtilities.OpenFile(GetPath());
        private void RunFileToolStripMenuItem_Click(object sender, EventArgs e) => WindowsUtilities.OpenFile(GetPath());
        private void RemoveAllDeletedFilesToolStripMenuItem_Click(object sender, EventArgs e) => RemoveAllDeleted();
        private void RemoveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var list = GetSelection();
            if (list.Count == 0)
                return;

            if (list.Count == 1)
            {
                if (this.ShowConfirm($"Are you sure you want to delete the shortcut to {list[0].Item.SIGDN_DESKTOPABSOLUTEPARSING}?") != DialogResult.Yes)
                    return;
            }
            else
            {
                if (this.ShowConfirm($"Are you sure you want to delete {list.Count} shortcuts?") != DialogResult.Yes)
                    return;
            }

            var count = 0;
            foreach (var group in list.GroupBy(i => i.AppUserModelID))
            {
                count += AutomaticDestinationList.RemoveItems(group.Key, group.Select(g => g.Item));
            }
            this.ShowMessage($"{count} shortcut(s) were deleted.");

            if (count > 0)
            {
                LoadApplicationUserModelIDs();
            }
        }

        private void ContextMenuStripMain_Opening(object sender, System.ComponentModel.CancelEventArgs e)
        {
            runFileToolStripMenuItem.Text = "Run File";
            runFileToolStripMenuItem.Enabled = false;
            openLocationToolStripMenuItem.Text = "Open Location";
            openLocationToolStripMenuItem.Enabled = false;

            var selection = GetSelection();
            if (selection.Count > 1)
            {
                removeToolStripMenuItem.Text = $"Remove {selection.Count} shortcuts";
            }
            else
            {
                var path = GetPath();
                if (path != null)
                {
                    runFileToolStripMenuItem.Text = $"Run '{path}'";
                    runFileToolStripMenuItem.Enabled = true;
                    removeToolStripMenuItem.Text = $"Remove '{path}' shortcut";
                }
            }

            var parentPath = GetParentPath();
            if (parentPath != null)
            {
                openLocationToolStripMenuItem.Text = $"Open '{parentPath}'";
                openLocationToolStripMenuItem.Enabled = true;
            }
        }

        private sealed class Model
        {
            public Model(string aumid, Item item)
            {
                AppUserModelID = aumid;
                Item = item;
            }

            public string AppUserModelID { get; }
            public Item Item { get; }
        }

        private sealed class ColumnSorter : IComparer
        {
            private static readonly CaseInsensitiveComparer _comparer = new();

            public int SortColumn { get; set; }
            public SortOrder Order { get; set; }

            public int Compare(object? x, object? y)
            {
                var listviewX = (ListViewItem)x!;
                var listviewY = (ListViewItem)y!;

                if (SortColumn >= listviewX.SubItems.Count)
                    return -1;

                if (SortColumn >= listviewY.SubItems.Count)
                    return 1;

                var compareResult = _comparer.Compare(listviewX.SubItems[SortColumn].Text, listviewY.SubItems[SortColumn].Text);
                if (Order == SortOrder.Ascending)
                    return compareResult;

                if (Order == SortOrder.Descending)
                    return -compareResult;

                return 0;
            }

            public void HandleClick(ListView listView, int column)
            {
                if (column == SortColumn)
                {
                    if (Order == SortOrder.Ascending)
                    {
                        Order = SortOrder.Descending;
                    }
                    else
                    {
                        Order = SortOrder.Ascending;
                    }
                }
                else
                {
                    SortColumn = column;
                    Order = SortOrder.Ascending;
                }
                listView.Sort();
            }
        }
    }
}
