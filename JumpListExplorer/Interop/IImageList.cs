using System;
using System.Runtime.InteropServices;

namespace JumpListExplorer.Interop
{
    [ComImport, Guid("46eb5926-582e-4017-9fdf-e8998daa0950"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal partial interface IImageList
    {
#pragma warning disable IDE1006 // Naming Styles
        void _VtblGap1_7(); // skip 7 methods we don't need
#pragma warning restore IDE1006 // Naming Styles

        [PreserveSig]
        int GetIcon(int i, int flags, out IntPtr picon);

        // the rest is undefined
    }
}
