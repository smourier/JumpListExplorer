using System.Reflection;
using System.Windows.Forms;

namespace JumpListExplorer.Utilities
{
    public static class WinformsUtilities
    {
        public static string ApplicationName => Resources.AppName;
        public static string? ApplicationVersion => AssemblyUtilities.GetFileVersion(Assembly.GetExecutingAssembly());
        public static string ApplicationTitle => ApplicationName + " V" + ApplicationVersion;

        public static void ShowMessage(this IWin32Window owner, string text) => MessageBox.Show(owner, text, ApplicationTitle, MessageBoxButtons.OK, MessageBoxIcon.Information);
        public static DialogResult ShowConfirm(this IWin32Window owner, string text) => MessageBox.Show(owner, text, ApplicationTitle + " - " + Resources.Confirmation, MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
        public static DialogResult ShowQuestion(this IWin32Window owner, string text) => MessageBox.Show(owner, text, ApplicationTitle + " - " + Resources.Confirmation, MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);
        public static void ShowError(this IWin32Window owner, string text) => MessageBox.Show(owner, text, ApplicationTitle, MessageBoxButtons.OK, MessageBoxIcon.Error);
        public static void ShowWarning(this IWin32Window owner, string text) => MessageBox.Show(owner, text, ApplicationTitle + " - " + Resources.Warning, MessageBoxButtons.OK, MessageBoxIcon.Warning);
    }
}
