using System;
using System.Collections.Generic;
using System.Globalization;
using Siemens.Engineering;
using Siemens.Engineering.AddIn.Menu;
using System.Linq;
using System.Windows.Forms;

namespace MultiArchiver
{
    public class AddIn : ContextMenuAddIn
    {
        private readonly TiaPortal _tiaPortal;
        private readonly Settings _settings;

        public AddIn(TiaPortal tiaPortal) : base("Example Add-In")
        {
            _tiaPortal = tiaPortal;
            _settings = Settings.Load();
        }

        protected override void BuildContextMenuItems(ContextMenuAddInRoot addInRootSubmenu)
        {
            addInRootSubmenu.Items.AddActionItem<IEngineeringObject>("Selection info", OnClick, DisplayStatus);
            Submenu settingsSubmenu = addInRootSubmenu.Items.AddSubmenu("Settings");
            settingsSubmenu.Items.AddActionItemWithCheckBox<IEngineeringObject>("Check Box", _settings.CheckBoxOnClick, _settings.CheckBoxDisplayStatus);
            settingsSubmenu.Items.AddActionItemWithRadioButton<IEngineeringObject>("Radio Button 1", _settings.RadioButton1OnClick, _settings.RadioButton1DisplayStatus);
            settingsSubmenu.Items.AddActionItemWithRadioButton<IEngineeringObject>("Radio Button 2", _settings.RadioButton2OnClick, _settings.RadioButton2DisplayStatus);
        }

        private void OnClick(MenuSelectionProvider<IEngineeringObject> menuSelectionProvider)
        {
            // TODO: Replace this with your own on click logic

            string projectName = _tiaPortal.Projects.First(project => project.IsPrimary).Name;
            List<IEngineeringObject> selectedObjects = menuSelectionProvider.GetSelection<IEngineeringObject>().ToList();
            string selectedObjectNames = string.Join(Environment.NewLine, selectedObjects.Select(selection => (string) selection.GetAttribute("Name")));

            string message = string.Format(CultureInfo.InvariantCulture, "Project name: {0}\r\nSelection:\r\n{1}", projectName, selectedObjectNames);
            string title = "TIA Add-In: Selection info";
            using (Form owner = Util.GetForegroundWindow())
            {
                MessageBox.Show(owner, message, title);
            }
        }

        private MenuStatus DisplayStatus(MenuSelectionProvider<IEngineeringObject> menuSelectionProvider)
        {
            return MenuStatus.Enabled;
        }
    }
}
