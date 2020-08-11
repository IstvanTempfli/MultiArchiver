using Siemens.Engineering;
using Siemens.Engineering.AddIn.Menu;
using System;
using System.IO;
using System.Xml.Serialization;

namespace MultiArchiver
{
    public class Settings
    {
        public bool CheckBox { get; set; }
        public int RadioButton { get; set; }

        private static readonly string SettingsFilePath;

        static Settings()
        {
            string settingsDirectory = Path.Combine(Environment.ExpandEnvironmentVariables("%TEMP%"), AppDomain.CurrentDomain.FriendlyName);
            Directory.CreateDirectory(settingsDirectory);
            SettingsFilePath = Path.Combine(settingsDirectory, "Settings.xml");
        }

        public Settings()
        {
            CheckBox = true;
            RadioButton = 1;
        }

        public static Settings Load()
        {
            if (File.Exists(SettingsFilePath) == false)
            {
                return new Settings();
            }

            using (FileStream readStream = new FileStream(SettingsFilePath, FileMode.Open))
            {
                XmlSerializer serializer = new XmlSerializer(typeof(Settings));
                return serializer.Deserialize(readStream) as Settings;
            }
        }

        public void Save()
        {
            using (FileStream writeStream = new FileStream(SettingsFilePath, FileMode.Create))
            {
                XmlSerializer serializer = new XmlSerializer(typeof(Settings));
                serializer.Serialize(writeStream, this);
            }
        }
        
        internal void CheckBoxOnClick(MenuSelectionProvider<IEngineeringObject> menuSelectionProvider)
        {
            // TODO: Replace this with your own check box logic
            CheckBox = !CheckBox;
            Save();
        }
        internal MenuStatus CheckBoxDisplayStatus(MenuSelectionProvider<IEngineeringObject> menuSelectionProvider, CheckBoxActionItemStyle checkBoxStyle)
        {
            checkBoxStyle.State = CheckBox == true ? CheckBoxState.Checked : CheckBoxState.Unchecked;
            return MenuStatus.Enabled;
        }

        internal void RadioButton1OnClick(MenuSelectionProvider<IEngineeringObject> menuSelectionProvider)
        {
            // TODO: Replace this with your own radio button logic
            RadioButton = 1;
            Save();
        }
        internal MenuStatus RadioButton1DisplayStatus(MenuSelectionProvider<IEngineeringObject> menuSelectionProvider, RadioButtonActionItemStyle radioButtonStyle)
        {
            radioButtonStyle.State = RadioButton == 1 ? RadioButtonState.Selected : RadioButtonState.Unselected;
            return MenuStatus.Enabled;
        }
        
        internal void RadioButton2OnClick(MenuSelectionProvider<IEngineeringObject> menuSelectionProvider)
        {
            // TODO: Replace this with your own radio button logic
            RadioButton = 2;
            Save();
        }
        internal MenuStatus RadioButton2DisplayStatus(MenuSelectionProvider<IEngineeringObject> menuSelectionProvider, RadioButtonActionItemStyle radioButtonStyle)
        {
            radioButtonStyle.State = RadioButton == 2 ? RadioButtonState.Selected : RadioButtonState.Unselected;
            return MenuStatus.Enabled;
        }
    }
}
