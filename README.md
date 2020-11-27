# MultiArchiver
TIA Portal Add-In for archiving project to multiple folders
# Installation

Please copy the .addin-file into the "AddIns" folder of your TIA Portal installation path.
Per default, it is C:\Program Files\Siemens\Automation\Portal V16\AddIns\

# Features
* **Archive Project** - Arcives the current project to the saved paths.

* **View Folders** - Display in a dialog the target folders with comment "Ok" or "Not found"

# Settings
* **Edit Folders** - Opens, using the operating systems default _*.txt_ editor (e.g. Edit) the _FolderList.txt_ file. Simply copy and paste the required paths into this file,
  save and close normally. The file is saved into the project files, thus each TIA Project will have its own archive target list.
* **Move old files to the Archive folder** - Automatically moves the old archive files into an "Archive" folder in all target paths.
* **Show not found folders when completed** - Displays a dialog at the end, listing the not reacheable folders.
* **Debug** - Used for verbose logging
