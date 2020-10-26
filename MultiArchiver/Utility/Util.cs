using System;
using System.IO;
using System.Windows.Forms;
using System.Collections.Generic;
using Siemens.Engineering;
using Siemens.Engineering.AddIn.Menu;
using System.Linq;
using System.Globalization;
using System.Reflection;

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
    }
}
