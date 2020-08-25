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
    public class Settings
    {
        public bool MoveOldFiles { get; set; }
        public bool Debug { get; set; }

        private static string settingsFilePath;

        static Settings()
        {

        }

        public Settings()
        {
            MoveOldFiles = true;
            Debug = true;
        }

        public static Settings Load(DirectoryInfo settingsDirectory)
        {
            settingsFilePath = Path.Combine(settingsDirectory.FullName, string.Concat(typeof(Settings).Name, ".xml"));

            if (File.Exists(settingsFilePath) == false)
            {
                return new Settings();
            }

            using (FileStream readStream = new FileStream(settingsFilePath, FileMode.Open))
            {
                XmlSerializer serializer = new XmlSerializer(typeof(Settings));
                return serializer.Deserialize(readStream) as Settings;
            }
        }

        public void Save()
        {
            using (FileStream writeStream = new FileStream(settingsFilePath, FileMode.Create))
            {
                XmlSerializer serializer = new XmlSerializer(typeof(Settings));
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

    }
}
