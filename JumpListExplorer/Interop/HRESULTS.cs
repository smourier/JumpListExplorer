namespace JumpListExplorer.Interop
{
    public enum HRESULTS : uint
    {
        S_OK = 0,
        S_FALSE = 1,
        E_FAIL = 0x80004005,
        E_NOINTERFACE = 0x80004002,
        E_INVALIDARG = 0x80070057,
        E_ACCESSDENIED = 0x80070005,
        E_NOTIMPL = 0x80004001,
        E_UNEXPECTED = 0x8000FFFF,
        DISP_E_EXCEPTION = 0x80020009,
        STG_E_FILENOTFOUND = 0x80030002,
        STG_E_PATHNOTFOUND = 0x80030003,
    }
}
