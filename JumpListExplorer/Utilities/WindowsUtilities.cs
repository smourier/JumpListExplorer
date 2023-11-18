using System.Diagnostics;
using System.IO;

namespace JumpListExplorer.Utilities
{
    public static class WindowsUtilities
    {
        public static void OpenFile(string? fileName)
        {
            if (fileName == null)
                return;

            var psi = new ProcessStartInfo
            {
                FileName = fileName,
                UseShellExecute = true
            };
            Process.Start(psi);
        }

        public static void OpenExplorer(string? directoryPath)
        {
            if (directoryPath == null)
                return;

            if (!Directory.Exists(directoryPath))
                return;

            // see http://support.microsoft.com/kb/152457/en-us
            Process.Start("explorer.exe", "/e,/root,/select," + directoryPath);
        }
    }
}
