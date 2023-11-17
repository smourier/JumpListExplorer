using System;
using System.Drawing;
using JumpListExplorer.Interop;
using JumpListExplorer.Utilities;

namespace JumpListExplorer.Shell
{
    public class Item
    {
        internal readonly IShellItem _shellItem;
        internal readonly IShellItem2 _shellItem2;

        public Item(object? shellItem)
        {
            var item = shellItem as IShellItem;
            ArgumentNullException.ThrowIfNull(item, nameof(shellItem));
            _shellItem = item;
            _shellItem2 = (IShellItem2)shellItem!;
        }

        public object NativeObject => _shellItem;
        public string SIGDN_DESKTOPABSOLUTEEDITING { get { _shellItem.GetDisplayName(SIGDN.SIGDN_DESKTOPABSOLUTEEDITING, out var name); return name; } }
        public string SIGDN_DESKTOPABSOLUTEPARSING { get { _shellItem.GetDisplayName(SIGDN.SIGDN_DESKTOPABSOLUTEPARSING, out var name); return name; } }
        public string SIGDN_FILESYSPATH { get { _shellItem.GetDisplayName(SIGDN.SIGDN_FILESYSPATH, out var name); return name; } }
        public string SIGDN_NORMALDISPLAY { get { _shellItem.GetDisplayName(SIGDN.SIGDN_NORMALDISPLAY, out var name); return name; } }
        public string SIGDN_PARENTRELATIVE { get { _shellItem.GetDisplayName(SIGDN.SIGDN_PARENTRELATIVE, out var name); return name; } }
        public string SIGDN_PARENTRELATIVEEDITING { get { _shellItem.GetDisplayName(SIGDN.SIGDN_PARENTRELATIVEEDITING, out var name); return name; } }
        public string SIGDN_PARENTRELATIVEFORADDRESSBAR { get { _shellItem.GetDisplayName(SIGDN.SIGDN_PARENTRELATIVEFORADDRESSBAR, out var name); return name; } }
        public string SIGDN_PARENTRELATIVEFORUI { get { _shellItem.GetDisplayName(SIGDN.SIGDN_PARENTRELATIVEFORUI, out var name); return name; } }
        public string SIGDN_PARENTRELATIVEPARSING { get { _shellItem.GetDisplayName(SIGDN.SIGDN_PARENTRELATIVEPARSING, out var name); return name; } }
        public string SIGDN_URL { get { _shellItem.GetDisplayName(SIGDN.SIGDN_URL, out var name); return name; } }
        public SFGAO Attributes { get { _shellItem.GetAttributes((SFGAO)0x7FFFFFFF, out var flags); return flags; } }
        public bool IsFolder => Attributes.HasFlag(SFGAO.SFGAO_FOLDER);
        public bool IsHidden => Attributes.HasFlag(SFGAO.SFGAO_HIDDEN);
        public bool IsReadOnly => Attributes.HasFlag(SFGAO.SFGAO_READONLY);
        public Folder? Parent { get { _shellItem.GetParent(out var item); return item == null ? null : new Folder(item); } }
        public long? Size => GetProperty<long?>(PROPERTYKEY.System.Size);
        public DateTime? DateModified => GetProperty<DateTime?>(PROPERTYKEY.System.DateModified);
        public DateTime? DateAccessed => GetProperty<DateTime?>(PROPERTYKEY.System.DateAccessed);
        public DateTime? DateCreated => GetProperty<DateTime?>(PROPERTYKEY.System.DateCreated);
        public string? ItemType => GetProperty<string>(PROPERTYKEY.System.ItemType);
        public string? ItemTypeText => GetProperty<string>(PROPERTYKEY.System.ItemTypeText);

        public object? GetProperty(PROPERTYKEY pk, bool throwOnError = false)
        {
            using var pv = new PROPVARIANT();
            _shellItem2.GetProperty(pk, pv).ThrowOnError(throwOnError);
            return pv.Value;
        }

        public T? GetProperty<T>(PROPERTYKEY pk, T? defaultValue = default)
        {
            using var pv = new PROPVARIANT();
            var hr = _shellItem2.GetProperty(pk, pv);
            if (hr.IsError)
                return defaultValue;

            var value = pv.Value;
            if (value is T t)
                return t;

            return Conversions.ChangeType(value, defaultValue);
        }

        public override string ToString() => SIGDN_NORMALDISPLAY;

        public IntPtr GetIconHandleFromImageList(SHIL shil) => IconUtilities.GetIconHandleFromImageList(_shellItem, shil);
        public Icon? GetIconFromImageList(SHIL shil)
        {
            var handle = GetIconHandleFromImageList(shil);
            if (handle == IntPtr.Zero)
                return null;

            using var icon = Icon.FromHandle(handle);
            var clone = (Icon)icon.Clone();
            Native.DestroyIcon(handle);
            return clone;
        }

        internal static Item? ToItem(IShellItem? shellItem)
        {
            if (shellItem == null)
                return null;

            shellItem.GetAttributes(SFGAO.SFGAO_FOLDER, out var flags);
            if (flags.HasFlag(SFGAO.SFGAO_FOLDER))
                return new Folder(shellItem);

            return new Item(shellItem);
        }
    }
}
