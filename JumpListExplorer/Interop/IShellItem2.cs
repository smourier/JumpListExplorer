using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace JumpListExplorer.Interop
{
    [ComImport, Guid("7e9fb0d3-919f-4307-ab2e-9b1860310c93"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IShellItem2 : IShellItem
    {
        // IShellItem
        [PreserveSig]
        new HRESULT BindToHandler(IBindCtx pbc, [MarshalAs(UnmanagedType.LPStruct)] Guid bhid, [MarshalAs(UnmanagedType.LPStruct)] Guid riid, [MarshalAs(UnmanagedType.IUnknown)] out object ppv);

        [PreserveSig]
        new HRESULT GetParent(out IShellItem ppsi);

        [PreserveSig]
        new HRESULT GetDisplayName(SIGDN sigdnName, [MarshalAs(UnmanagedType.LPWStr)] out string ppszName);

        [PreserveSig]
        new HRESULT GetAttributes(SFGAO sfgaoMask, out SFGAO psfgaoAttribs);

        [PreserveSig]
        new HRESULT Compare(IShellItem psi, SICHINTF hint, out int piOrder);

        // IShellItem2
        [PreserveSig]
        HRESULT GetPropertyStore(GETPROPERTYSTOREFLAGS flags, [MarshalAs(UnmanagedType.LPStruct)] Guid riid, [MarshalAs(UnmanagedType.IUnknown)] out object ppv);

        [PreserveSig]
        HRESULT GetPropertyStoreWithCreateObject(GETPROPERTYSTOREFLAGS flags, [MarshalAs(UnmanagedType.IUnknown)] object punkCreateObject, [MarshalAs(UnmanagedType.LPStruct)] Guid riid, [MarshalAs(UnmanagedType.IUnknown)] out object ppv);

        [PreserveSig]
        HRESULT GetPropertyStoreForKeys([MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 1)] PROPERTYKEY[] rgKeys, int cKeys, GETPROPERTYSTOREFLAGS flags, [MarshalAs(UnmanagedType.LPStruct)] Guid riid, [MarshalAs(UnmanagedType.IUnknown)] out object ppv);

        [PreserveSig]
        HRESULT GetPropertyDescriptionList(ref PROPERTYKEY keyType, [MarshalAs(UnmanagedType.LPStruct)] Guid riid, [MarshalAs(UnmanagedType.IUnknown)] out object ppv);

        [PreserveSig]
        HRESULT Update(IBindCtx pbc);

        [PreserveSig]
        HRESULT GetProperty(ref PROPERTYKEY key, [In, Out] PROPVARIANT ppropvar);

        [PreserveSig]
        HRESULT GetCLSID(ref PROPERTYKEY key, out Guid pclsid);

        [PreserveSig]
        HRESULT GetFileTime(ref PROPERTYKEY key, out long pft);

        [PreserveSig]
        HRESULT GetInt32(ref PROPERTYKEY key, out int pi);

        [PreserveSig]
        HRESULT GetString(ref PROPERTYKEY key, out IntPtr ppsz);

        [PreserveSig]
        HRESULT GetUInt32(ref PROPERTYKEY key, out uint pui);

        [PreserveSig]
        HRESULT GetUInt64(ref PROPERTYKEY key, out ulong pull);

        [PreserveSig]
        HRESULT GetBool(ref PROPERTYKEY key, out bool pf);
    }
}
