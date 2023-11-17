using System;
using System.Runtime.InteropServices;

namespace JumpListExplorer.Interop
{
    [ComImport, Guid("3acf075f-71db-4afa-81f0-3fc4fdf2a5b8"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal partial interface IFileSystemBindData2
    {
        void SetFindData(ref WIN32_FIND_DATAW pfd);
        void GetFindData(out WIN32_FIND_DATAW pfd);
        void SetFileID(long liFileID);
        void GetFileID(out long pliFileID);
        void SetJunctionCLSID([MarshalAs(UnmanagedType.LPStruct)] Guid clsid);
        void GetJunctionCLSID(out Guid pclsid);
    }
}
