using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace JumpListExplorer.Interop
{
    [ComImport, Guid("43826d1e-e718-42ee-bc55-a1e261c37bfe"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IShellItem
    {
        [PreserveSig]
        HRESULT BindToHandler(IBindCtx? pbc, [MarshalAs(UnmanagedType.LPStruct)] Guid bhid, [MarshalAs(UnmanagedType.LPStruct)] Guid riid, [MarshalAs(UnmanagedType.IUnknown)] out object ppv);

        [PreserveSig]
        HRESULT GetParent(out IShellItem ppsi);

        [PreserveSig]
        HRESULT GetDisplayName(SIGDN sigdnName, [MarshalAs(UnmanagedType.LPWStr)] out string ppszName);

        [PreserveSig]
        HRESULT GetAttributes(SFGAO sfgaoMask, out SFGAO psfgaoAttribs);

        [PreserveSig]
        HRESULT Compare(IShellItem psi, SICHINTF hint, out int piOrder);
    }
}
