using System;
using System.Runtime.InteropServices;

namespace JumpListExplorer.Interop
{
    [ComImport, Guid("0000010c-0000-0000-c000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal partial interface IPersist
    {
        [PreserveSig]
        HRESULT GetClassID(out Guid pClassID);
    }
}
