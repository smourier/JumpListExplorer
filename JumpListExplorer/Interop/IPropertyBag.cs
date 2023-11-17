using System;
using System.Runtime.InteropServices;

namespace JumpListExplorer.Interop
{
    [ComImport, Guid("55272A00-42CB-11CE-8135-00AA004BB851"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IPropertyBag
    {
        [PreserveSig]
        HRESULT Read([MarshalAs(UnmanagedType.LPWStr)] string pszPropName, out object pVar, IntPtr pErrorLog);

        [PreserveSig]
        HRESULT Write([MarshalAs(UnmanagedType.LPWStr)] string gpszPropName, ref object pVar);
    }
}
