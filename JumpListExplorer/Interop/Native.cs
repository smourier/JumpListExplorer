using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace JumpListExplorer.Interop
{
    internal static class Native
    {
        public static Guid IID_IUnknown = new Guid("00000000-0000-0000-c000-000000000046");

        [DllImport("shell32")]
        public static extern HRESULT SHGetImageList(SHIL iImageList, [MarshalAs(UnmanagedType.LPStruct)] Guid riid, out IImageList ppv);

        [DllImport("shell32")]
        public static extern int SHMapPIDLToSystemImageListIndex(IntPtr pshf, IntPtr pidl, out int piIndexSel);

        [DllImport("shell32")]
        public static extern HRESULT SHGetKnownFolderItem(Guid rfid, KNOWN_FOLDER_FLAG flags, IntPtr hToken, [MarshalAs(UnmanagedType.LPStruct)] Guid riid, [MarshalAs(UnmanagedType.IUnknown)] out object ppv);

        [DllImport("shell32")]
        public static extern HRESULT SHCreateItemFromParsingName([MarshalAs(UnmanagedType.LPWStr)] string pszPath, IBindCtx? pbc, [MarshalAs(UnmanagedType.LPStruct)] Guid riid, out IShellItem ppv);

        [DllImport("user32")]
        public static extern bool DestroyIcon(IntPtr handle);

        [DllImport("ole32")]
        private static extern HRESULT CreateBindCtx(int reserved, out IBindCtx ppbc);

        [DllImport("propsys")]
        public extern static HRESULT PSCreateMemoryPropertyStore([MarshalAs(UnmanagedType.LPStruct)] Guid riid, out IPropertyStore store);

        public static IBindCtx CreateBindCtx()
        {
            CreateBindCtx(0, out var ctx).ThrowOnError();
            return ctx;
        }
    }
}
