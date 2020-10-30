using System;
using System.IO;
using System.Windows.Forms;
using System.Collections.Generic;
using Siemens.Engineering;
using Siemens.Engineering.AddIn.Menu;
using System.Linq;
using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;

namespace MultiArchiver.Utility
{
    internal static class Util
    {

        internal static Form GetForegroundWindow()
        {
            // Workaround for Add-In Windows to be shown in foreground of TIA Portal
            Form form = new Form { Opacity = 0, ShowIcon = false };
            form.Show();
            form.TopMost = true;
            form.Activate();
            form.TopMost = false;
            return form;
        }

        /// <summary>
        /// Search for a specified text in the project comment and returns the comment after the phrase.
        /// </summary>
        /// <param name="tia">TiaPortal instance</param>
        /// <param name="searchPhrase">The search text</param>
        /// <returns>The text after the search phrase in string format</returns>
        public static string ReadProjectComment(TiaPortal tia, string searchPhrase)
        {
            Project project = tia.Projects.First();
            string returnStr = String.Empty;
            foreach (MultilingualTextItem comment in project.Comment.Items)
            {
                if (comment.Text.Contains(searchPhrase))
                {
                    returnStr = comment.Text.Substring(comment.Text.LastIndexOf(searchPhrase) + searchPhrase.Length);
                    break;
                }
            }
            return returnStr;
        }

        public static List<DirectoryInfo> LoadDirectories(string textPaths)
        {
            List<DirectoryInfo> savePaths = new List<DirectoryInfo>();

            var lines = textPaths.Split("\n\r".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);

            foreach (string path in lines)
            {
                savePaths.Add(new DirectoryInfo(path));
            }
            return savePaths;
        }

        //public static string GetPathList(string text)
        //{
        //    string paths = "";
        //    int i = 1;
        //    foreach (DirectoryInfo dir in LoadDirectories(text))
        //    {
        //        paths = paths + i.ToString() + ". " + dir.FullName + "\r\n";
        //        i++;
        //    }
        //    return paths;
        //}

        public static string GetPathList(List<DirectoryInfo> directories)
        {
            string paths = "";
            int i = 1;
            foreach (DirectoryInfo dir in directories)
            {
                //paths += i.ToString() + ". " + dir.FullName + Environment.NewLine;
                string exists = dir.Exists ? "Ok" : "Not Found";
                paths += dir.FullName + " - " + exists + Environment.NewLine;
                i++;
            }
            return paths;
        }

        private enum FO_Func : uint
        {
            FO_COPY = 0x0002,
            FO_DELETE = 0x0003,
            FO_MOVE = 0x0001,
            FO_RENAME = 0x0004,
        }

        private struct SHFILEOPSTRUCT
        {
            public IntPtr hwnd;
            public FO_Func wFunc;
            [MarshalAs(UnmanagedType.LPWStr)]
            public string pFrom;
            [MarshalAs(UnmanagedType.LPWStr)]
            public string pTo;
            public ushort fFlags;
            public bool fAnyOperationsAborted;
            public IntPtr hNameMappings;
            [MarshalAs(UnmanagedType.LPWStr)]
            public string lpszProgressTitle;

        }

        [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
        static extern int SHFileOperation([In] ref SHFILEOPSTRUCT
           lpFileOp);

        private static SHFILEOPSTRUCT _ShFile;

        public static void CopyFiles(string sSource, string sTarget)
        {

            try
            {
                _ShFile.wFunc = FO_Func.FO_COPY;
                _ShFile.pFrom = sSource + "\0\0";
                _ShFile.pTo = sTarget + "\0\0";
                SHFileOperation(ref _ShFile);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }
}
