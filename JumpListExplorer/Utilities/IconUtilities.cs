using System;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using JumpListExplorer.Interop;

namespace JumpListExplorer.Utilities
{
    public sealed class IconUtilities
    {
        public static Icon? GetIconFromImageList(string path, SHIL shil)
        {
            var handle = GetIconHandleFromImageList(path, shil);
            if (handle == IntPtr.Zero)
                return null;

            using var icon = Icon.FromHandle(handle);
            var clone = (Icon)icon.Clone();
            Native.DestroyIcon(handle);
            return clone;
        }

        // note: you must call DestroyIcon on the returned icon handle once you have finished to use it
        public static IntPtr GetIconHandleFromImageList(string path, SHIL shil)
        {
            ArgumentNullException.ThrowIfNull(path);
            var ctx = CreateBindCtx(path);
            _ = Native.SHCreateItemFromParsingName(path, ctx, typeof(IShellItem).GUID, out var obj);
            if (obj == null)
                return IntPtr.Zero;

            return GetIconHandleFromImageList(obj, shil);
        }

        // create an IBindCtx for an item that doesn't exist
        private static IBindCtx CreateBindCtx(string name)
        {
            var data = new WIN32_FIND_DATAW
            {
                cFileName = Path.GetFileName(name),
            };

            var bindData = new FileSystemBindData2();
            bindData.SetFindData(ref data);
            var ctx = Native.CreateBindCtx();

            const string STR_FILE_SYS_BIND_DATA = "File System Bind Data";
            try
            {
                ctx.RegisterObjectParam(STR_FILE_SYS_BIND_DATA, bindData);
            }
            catch
            {
                // do nothing
            }

            var opts = new BIND_OPTS
            {
                cbStruct = Marshal.SizeOf<BIND_OPTS>(),
                grfMode = STGM_CREATE
            };
            ctx.SetBindOptions(ref opts);
            try
            {
                ctx.SetBindOptions(ref opts);
            }
            catch
            {
                // continue
            }

            return ctx;
        }

        internal static IntPtr GetIconHandleFromImageList(IShellItem item, SHIL shil)
        {
            if (item is not IParentAndItem pai)
                return IntPtr.Zero;

            if (pai.GetParentAndItem(IntPtr.Zero, out var psf, out var child) < 0)
                return IntPtr.Zero;

            var index = Native.SHMapPIDLToSystemImageListIndex(psf, child, out _);
            Marshal.FreeCoTaskMem(child);

            var list = GetImageList(shil);
            if (list == null)
                return IntPtr.Zero;

            list.GetIcon(index, 0, out var hicon);
            return hicon;
        }

        private static IImageList GetImageList(SHIL shil) { _ = Native.SHGetImageList(shil, typeof(IImageList).GUID, out var ilist); return ilist; }

        private const int STGM_CREATE = 0x00001000;

        private sealed class FileSystemBindData2 : IFileSystemBindData2
        {
            private Guid _clsid;
            private long _fileId;
            private WIN32_FIND_DATAW _data;

            public void SetFindData(ref WIN32_FIND_DATAW data) { _data = data; }
            public void GetFindData(out WIN32_FIND_DATAW data) { data = _data; }
            public void SetFileID(long fileID) { _fileId = fileID; }
            public void GetFileID(out long fileId) { fileId = _fileId; }
            public void SetJunctionCLSID(Guid clsid) { _clsid = clsid; }
            public void GetJunctionCLSID(out Guid clsid) { clsid = _clsid; }
        }
    }
}
