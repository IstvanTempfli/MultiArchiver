using System;
using System.Collections.Generic;
using System.Globalization;
using System.Reflection;
using System.IO;
using Siemens.Engineering;
using Siemens.Engineering.AddIn.Menu;
using Siemens.Engineering.AddIn.Utilities;
using System.Linq;
using System.Windows.Forms;
using MultiArchiver.Utility;

namespace MultiArchiver
{
    public class AddIn : ContextMenuAddIn
    {
        private readonly TiaPortal _tiaPortal;
        private readonly AddinSettings _addinSettings;

        private string projectName = String.Empty;

        private readonly string _traceFilePath;

        public AddIn(TiaPortal tiaPortal) : base("MultiArchiver")
        {
            _tiaPortal = tiaPortal;

            var assemblyName = Assembly.GetCallingAssembly().GetName();
            var logDirectoryPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "TIA Add-Ins", assemblyName.Name, assemblyName.Version.ToString(), "Logs");
            var logDirectory = Directory.CreateDirectory(logDirectoryPath);
            _traceFilePath = Path.Combine(logDirectory.FullName, string.Concat(DateTime.Now.ToString("yyyyMMdd"), ".txt"));

            var settingsDirectoryPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "TIA Add-Ins", assemblyName.Name, assemblyName.Version.ToString());
            var settingsDirectory = Directory.CreateDirectory(settingsDirectoryPath);

            _addinSettings = AddinSettings.Load(settingsDirectory);

            WriteLog("Add-in started");
        }

        protected override void BuildContextMenuItems(ContextMenuAddInRoot addInRootSubmenu)
        {

            WriteLog("Building context menu");

            addInRootSubmenu.Items.AddActionItem<Project>("Archive Project", ArchiveOnClick); //Main funtion
            addInRootSubmenu.Items.AddActionItem<IEngineeringObject>("Please select the project.", menuSelectionProvider => { }, InfoTextStatus);
            addInRootSubmenu.Items.AddActionItem<IEngineeringObject>("View Folders", ViewFoldersOnClick);

            Submenu settingsSubmenu = addInRootSubmenu.Items.AddSubmenu("Settings"); //Settings submenu
            settingsSubmenu.Items.AddActionItem<IEngineeringObject>("Edit Folders", EditFoldersOnClick);
            settingsSubmenu.Items.AddActionItemWithCheckBox<IEngineeringObject>("Move old files to the Archive folder", _addinSettings.MoveOldFilesOnClick, _addinSettings.MoveOldFilesDisplayStatus);
            settingsSubmenu.Items.AddActionItemWithCheckBox<IEngineeringObject>("Show not found folders when completed", _addinSettings.ShowSummaryOnClick, _addinSettings.ShowSummaryDisplayStatus);
            settingsSubmenu.Items.AddActionItemWithCheckBox<IEngineeringObject>("Debug mode", _addinSettings.DebugOnClick, _addinSettings.DebugDisplayStatus);
        }

        private void ArchiveOnClick(MenuSelectionProvider menuSelectionProvider)
        {
            List<DirectoryInfo> paths, pathsDone, pathsError;

            paths = Util.LoadDirectoriesFromFile(Util.GetFolderListPath(_tiaPortal));

            pathsDone = new List<DirectoryInfo>();
            pathsError = new List<DirectoryInfo>();

            Project project = _tiaPortal.Projects.First();
            projectName = project.Name;

            string archiveName = String.Format("{0}_{1}.zap16", project.Name, DateTime.Now.ToString("yyyyMMdd_HHmm"));

            //Save Project
            project.Save();
            WriteLog("Project Saved");

            using (var exclusiveAccess = _tiaPortal.ExclusiveAccess("Archiving to " + paths.Count + " folders..."))
            {
                WriteLog("Archiving to " + paths.Count + " folders");
                using (var transaction = exclusiveAccess.Transaction(project, "Archive project"))
                {

                    var showMessageBox = false;
                    string sourceFile = "", targetFile;

                    foreach (DirectoryInfo path in paths)
                    {
                        if (path.Exists)
                        {
                            //Paths exists, begin arcive
                            if (_addinSettings.MoveOldFiles)
                            {
                                ArchiveFiles(project.Name, path, "Archiv");
                            }

                            if (pathsDone.Count == 0)
                            {
                                exclusiveAccess.Text = "Archiving to: " + path.FullName;
                                //Archive to the first path
                                project.Archive(path, archiveName, ProjectArchivationMode.DiscardRestorableDataAndCompressed);
                                sourceFile = Path.Combine(path.FullName, archiveName);
                                WriteLog("Archive: " + sourceFile);
                            }
                            else
                            {
                                exclusiveAccess.Text = "Copy archive to: " + path.FullName;
                                //Copy from the first archive
                                targetFile = Path.Combine(path.FullName, archiveName);
                                Util.CopyFiles(sourceFile, targetFile);
                                WriteLog("Copy to: " + targetFile);
                            }

                            pathsDone.Add(path);
                        }
                        else
                        {
                            //not found/incorrect
                            showMessageBox = _addinSettings.DisplaySummary;
                            pathsError.Add(path);
                            WriteLog("Not found: " + path.FullName);
                        }
                    }

                    if (showMessageBox)
                    {
                        exclusiveAccess.Text = "Completed!" + Environment.NewLine + "See the message box for further information.";
                        using (var owner = Util.GetForegroundWindow())
                        {
                            MessageBox.Show(owner, "The following paths could not be found:" + Environment.NewLine + Environment.NewLine + Util.PrintPathList(pathsError, false), "MultiArchiver", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }

                    if (transaction.CanCommit)
                    {
                        transaction.CommitOnDispose();
                    }
                }
            }
        }

        private void ArchiveFiles(string projectName, DirectoryInfo path, string archiveFolderName)
        {
            //Get list of older archives
            string[] files = Directory.GetFiles(path.FullName, projectName + "*.zap16");

            if (files.Length > 0)
            {

                try
                {
                    string target = Path.Combine(path.FullName, archiveFolderName);
                    WriteLog("Archive - Target Dir: " + target);

                    //Create subdirectory
                    if (!Directory.Exists(target))
                    {
                        Directory.CreateDirectory(target);
                        WriteLog("Archive - Archive dir created");
                    }

                    foreach (string file in files)
                    {
                        Util.MoveFiles(file, Path.Combine(target, Path.GetFileName(file)));
                        WriteLog("Moving " + file + " to: " + target);
                    }
                }
                catch (Exception e)
                {
                    WriteLog("Exception: " + e.ToString());
                }
            }
        }

        private void ViewFoldersOnClick(MenuSelectionProvider menuSelectionProvider)
        {
            string message = string.Format(CultureInfo.InvariantCulture, "Archive targets:\r\n{0}", Util.PrintPathList(Util.LoadDirectoriesFromFile(Util.GetFolderListPath(_tiaPortal)), true));
            string title = "MultiArchiver: Paths info";
            using (Form owner = Util.GetForegroundWindow())
            {
                MessageBox.Show(owner, message, title);
            }
        }

        private void EditFoldersOnClick(MenuSelectionProvider menuSelectionProvider)
        {
            try
            {
                string path = Util.GetFolderListPath(_tiaPortal);

                Process.Start(path);
            }
            catch (Exception e)
            {
                WriteLog("Exception: " + e.ToString());
            }
        }

        private static MenuStatus InfoTextStatus(MenuSelectionProvider<IEngineeringObject> menuSelectionProvider)
        {
            var show = false;

            foreach (IEngineeringObject engineeringObject in menuSelectionProvider.GetSelection())
            {
                if (!(engineeringObject.GetType() == menuSelectionProvider.GetSelection().First().GetType() && engineeringObject is Project))
                {
                    show = true;
                    break;
                }
            }
            return show ? MenuStatus.Disabled : MenuStatus.Hidden;
        }

        public void WriteLog(string text)
        {
            if (_addinSettings.Debug)
            {
                using (StreamWriter writer = new StreamWriter(new FileStream(_traceFilePath, FileMode.Append)))
                {
                    writer.WriteLine(projectName == String.Empty ? "{0}: {1}" : "{0} - {2}: {1}", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), text, projectName);
                }
            }
        }

    }
}
