using System;
using System.Runtime.InteropServices;

namespace JumpListExplorer.Interop
{
    [ComImport, Guid("b3a4b685-b685-4805-99d9-5dead2873236"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal partial interface IParentAndItem
    {
#pragma warning disable IDE1006 // Naming Styles
        void _VtblGap1_1(); // skip 1 methods we don't need
#pragma warning restore IDE1006 // Naming Styles

        [PreserveSig]
        int GetParentAndItem(IntPtr ppidlParent, out IntPtr ppsf, out IntPtr ppidlChild);
    }
}
