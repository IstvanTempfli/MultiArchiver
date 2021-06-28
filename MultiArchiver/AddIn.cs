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

        enum DebugLevel
        {
            Info,
            Warning,
            Error
        }

        public AddIn(TiaPortal tiaPortal) : base("MultiArchiver")
        {
            _tiaPortal = tiaPortal;

            try
            {
                var assemblyName = Assembly.GetCallingAssembly().GetName();
                var logDirectoryPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "TIA Add-Ins", assemblyName.Name, assemblyName.Version.ToString(), "Logs");
                var logDirectory = Directory.CreateDirectory(logDirectoryPath);
                _traceFilePath = Path.Combine(logDirectory.FullName, string.Concat(DateTime.Now.ToString("yyyyMMdd"), ".txt"));

                var settingsDirectoryPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "TIA Add-Ins", assemblyName.Name, assemblyName.Version.ToString());
                var settingsDirectory = Directory.CreateDirectory(settingsDirectoryPath);

                _addinSettings = AddinSettings.Load(settingsDirectory);
            }
            catch (Exception e)
            {
                WriteLog("Exception: " + e.ToString(), DebugLevel.Error);
            }
            //WriteLog("Add-in started", DebugLevel.Info);
        }

        protected override void BuildContextMenuItems(ContextMenuAddInRoot addInRootSubmenu)
        {
            WriteLog("Building context menu", DebugLevel.Info);
            try
            {
                addInRootSubmenu.Items.AddActionItem<IEngineeringObject>("Archive Project", ArchiveOnClick); //Main funtion
                //addInRootSubmenu.Items.AddActionItem<IEngineeringObject>("Please select the project.", menuSelectionProvider => { }, InfoTextStatus);
                addInRootSubmenu.Items.AddActionItem<IEngineeringObject>("View Folders", ViewFoldersOnClick);

                Submenu settingsSubmenu = addInRootSubmenu.Items.AddSubmenu("Settings"); //Settings submenu
                settingsSubmenu.Items.AddActionItem<IEngineeringObject>("Edit Folders", EditFoldersOnClick);
                settingsSubmenu.Items.AddActionItemWithCheckBox<IEngineeringObject>("Move old files to the Archive folder", _addinSettings.MoveOldFilesOnClick, _addinSettings.MoveOldFilesDisplayStatus);
                settingsSubmenu.Items.AddActionItemWithCheckBox<IEngineeringObject>("Show not found folders when completed", _addinSettings.ShowSummaryOnClick, _addinSettings.ShowSummaryDisplayStatus);
                settingsSubmenu.Items.AddActionItemWithCheckBox<IEngineeringObject>("Debug mode", _addinSettings.DebugOnClick, _addinSettings.DebugDisplayStatus);
            }
            catch (Exception e)
            {
                WriteLog("Exception: " + e.ToString(), DebugLevel.Error);
            }
        }

        private void ArchiveOnClick(MenuSelectionProvider menuSelectionProvider)
        {
            List<DirectoryInfo> paths, pathsDone, pathsError;

            paths = Util.LoadDirectoriesFromFile(Util.GetFolderListPath(_tiaPortal));

            if (paths.Count == 0)
            {
                string message = "No folders found.\nGo to Settings and select Edit Folders.";
                string title = "MultiArchiver: Error";
                using (Form owner = Util.GetForegroundWindow())
                {
                    MessageBox.Show(owner, message, title);
                    WriteLog("No folders found", DebugLevel.Warning);
                }
                return;
            }

            pathsDone = new List<DirectoryInfo>();
            pathsError = new List<DirectoryInfo>();

            Project project = _tiaPortal.Projects.First();
            projectName = project.Name;

            string archiveName = String.Format("{0}_{1}.zap16", project.Name, DateTime.Now.ToString("yyyyMMdd_HHmm"));

            //Save Project
            project.Save();
            WriteLog("Project Saved", DebugLevel.Info);

            using (var exclusiveAccess = _tiaPortal.ExclusiveAccess("Archiving to " + paths.Count + " folders..."))
            {
                WriteLog("Archiving to " + paths.Count + " folders", DebugLevel.Info);
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
                                exclusiveAccess.Text += Environment.NewLine + "Move old files: " + path.FullName;
                                if (ArchiveFiles(project.Name, path, "Archiv") > 0)
                                {
                                    exclusiveAccess.Text += " - Done";
                                }
                            }

                            if (pathsDone.Count == 0)
                            {
                                exclusiveAccess.Text += Environment.NewLine + "Archive: " + path.FullName;
                                //Archive to the first path
                                project.Archive(path, archiveName, ProjectArchivationMode.DiscardRestorableDataAndCompressed);
                                sourceFile = Path.Combine(path.FullName, archiveName);
                                WriteLog("Archive: " + sourceFile, DebugLevel.Info);
                                exclusiveAccess.Text += " - Done";
                            }
                            else
                            {
                                exclusiveAccess.Text += Environment.NewLine + "Copy to: " + path.FullName;
                                //Copy from the first archive
                                targetFile = Path.Combine(path.FullName, archiveName);
                                Util.CopyFiles(sourceFile, targetFile);
                                WriteLog("Copy to: " + targetFile, DebugLevel.Info);
                                exclusiveAccess.Text += " - Done";
                            }

                            pathsDone.Add(path);
                        }
                        else
                        {
                            //not found/incorrect
                            exclusiveAccess.Text += Environment.NewLine + path.FullName + " - Not found";
                            showMessageBox = _addinSettings.DisplaySummary;
                            pathsError.Add(path);
                            WriteLog("Not found: " + path.FullName, DebugLevel.Warning);
                        }
                    }

                    if (showMessageBox)
                    {
                        exclusiveAccess.Text += Environment.NewLine + "Completed! See the message box for further information.";
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

        private int ArchiveFiles(string projectName, DirectoryInfo path, string archiveFolderName)
        {
            //Get list of older archives
            string[] files = Directory.GetFiles(path.FullName, projectName + "*.zap16");
            int nrFiles = 0;

            if (files.Length > 0)
            {

                try
                {
                    string target = Path.Combine(path.FullName, archiveFolderName);
                    WriteLog("Archive - Target Dir: " + target, DebugLevel.Info);

                    //Create subdirectory
                    if (!Directory.Exists(target))
                    {
                        Directory.CreateDirectory(target);
                        WriteLog("Archive - Archive dir created", DebugLevel.Info);
                    }

                    foreach (string file in files)
                    {
                        Util.MoveFiles(file, Path.Combine(target, Path.GetFileName(file)));
                        WriteLog("Moving " + file + " to: " + target, DebugLevel.Info);
                        nrFiles++;
                    }
                }
                catch (Exception e)
                {
                    WriteLog("Exception: " + e.ToString(), DebugLevel.Error);
                    nrFiles = -1;
                }
            }

            return nrFiles;
        }

        private void ViewFoldersOnClick(MenuSelectionProvider menuSelectionProvider)
        {
            List<DirectoryInfo> paths = Util.LoadDirectoriesFromFile(Util.GetFolderListPath(_tiaPortal));
            string message;

            if (paths.Count == 0)
            {
                message = "No folders found.\nGo to Settings and select Edit Folders.";
            }
            else
            {
                message = string.Format(CultureInfo.InvariantCulture, "Archive Folders:\r\n{0}", Util.PrintPathList(paths, true));
            }

            string title = "MultiArchiver";
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
                WriteLog("Exception: " + e.ToString(), DebugLevel.Error);
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

        private void WriteLog(string text, DebugLevel level)
        {
            if (_addinSettings.Debug || level == DebugLevel.Error)
            {
                using (StreamWriter writer = new StreamWriter(new FileStream(_traceFilePath, FileMode.Append)))
                {
                    writer.WriteLine(String.Format("{0} | {1} | {2} | {3}", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff"), level.ToString().ToUpper(), projectName, text));
                }
            }
        }

    }
}
