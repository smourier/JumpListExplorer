using System;
using System.Runtime.InteropServices;

namespace JumpListExplorer.Interop
{
    [ComImport, Guid("70629033-e363-4a28-a567-0db78006e6d7"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal partial interface IEnumShellItems
    {
        [PreserveSig]
        HRESULT Next(int celt, out IShellItem rgelt, out int pceltFetched);

        [PreserveSig]
        HRESULT Skip(int celt);

        [PreserveSig]
        HRESULT Reset();

        [PreserveSig]
        HRESULT Clone(out IEnumShellItems ppenum);
    }
}
