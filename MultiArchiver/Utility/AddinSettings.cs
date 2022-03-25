using Siemens.Engineering;
using Siemens.Engineering.AddIn.Menu;
using System;
using System.Reflection;
using System.Collections.Generic;
using System.IO;
using System.Xml.Serialization;

namespace MultiArchiver.Utility
{
    [XmlRoot]
    [XmlType]
    public class AddinSettings
    {
        public bool MoveOldFiles { get; set; }
        public bool Debug { get; set; }
        public bool DisplaySummary { get; set; }
        public string ArchiveDirName { get; set; }

        private static string settingsFilePath;

        static AddinSettings()
        {

        }

        public AddinSettings()
        {
            MoveOldFiles = true;
            Debug = false;
            DisplaySummary = false;
            ArchiveDirName = "Archive";
        }

        public static AddinSettings Load(DirectoryInfo settingsDirectory)
        {
            settingsFilePath = Path.Combine(settingsDirectory.FullName, string.Concat(typeof(AddinSettings).Name, ".xml"));

            if (File.Exists(settingsFilePath) == false)
            {
                return new AddinSettings();
            }

            using (FileStream readStream = new FileStream(settingsFilePath, FileMode.Open))
            {
                XmlSerializer serializer = new XmlSerializer(typeof(AddinSettings));
                return serializer.Deserialize(readStream) as AddinSettings;
            }
        }

        public void Save()
        {
            using (FileStream writeStream = new FileStream(settingsFilePath, FileMode.Create))
            {
                XmlSerializer serializer = new XmlSerializer(typeof(AddinSettings));
                serializer.Serialize(writeStream, this);
            }
        }

        internal void MoveOldFilesOnClick(MenuSelectionProvider<IEngineeringObject> menuSelectionProvider)
        {
            MoveOldFiles = !MoveOldFiles;
            Save();
        }
        internal MenuStatus MoveOldFilesDisplayStatus(MenuSelectionProvider<IEngineeringObject> menuSelectionProvider, CheckBoxActionItemStyle checkBoxStyle)
        {
            checkBoxStyle.State = MoveOldFiles == true ? CheckBoxState.Checked : CheckBoxState.Unchecked;
            return MenuStatus.Enabled;
        }

        internal void DebugOnClick(MenuSelectionProvider<IEngineeringObject> menuSelectionProvider)
        {
            Debug = !Debug;
            Save();
        }
        internal MenuStatus DebugDisplayStatus(MenuSelectionProvider<IEngineeringObject> menuSelectionProvider, CheckBoxActionItemStyle checkBoxStyle)
        {
            checkBoxStyle.State = Debug == true ? CheckBoxState.Checked : CheckBoxState.Unchecked;
            return MenuStatus.Enabled;
        }

        internal void ShowSummaryOnClick(MenuSelectionProvider<IEngineeringObject> menuSelectionProvider)
        {
            DisplaySummary = !DisplaySummary;
            Save();
        }
        internal MenuStatus ShowSummaryDisplayStatus(MenuSelectionProvider<IEngineeringObject> menuSelectionProvider, CheckBoxActionItemStyle checkBoxStyle)
        {
            checkBoxStyle.State = DisplaySummary == true ? CheckBoxState.Checked : CheckBoxState.Unchecked;
            return MenuStatus.Enabled;
        }

    }
}
