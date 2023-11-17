using System;
using System.Runtime.InteropServices;

namespace JumpListExplorer.Interop
{
    [ComImport, Guid("886d8eeb-8cf2-4446-8d02-cdba1dbdcf99"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal partial interface IPropertyStore
    {
        [PreserveSig]
        HRESULT GetCount(out int cProps);

        [PreserveSig]
        HRESULT GetAt(int iProp, out PROPERTYKEY pkey);

        [PreserveSig]
        HRESULT GetValue(ref PROPERTYKEY key, [Out] PROPVARIANT pv);

        [PreserveSig]
        HRESULT SetValue(ref PROPERTYKEY key, PROPVARIANT propvar);

        [PreserveSig]
        HRESULT Commit();
    }
}
