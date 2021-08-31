# **Overview**

"Outlook Monitor" is a tool for monitoring and tracing item and folder changes in MAPI stores. Monitoring takes place in real-time.
All MAPI stores, which are loaded in a given MAPI profile, can be monitored.

The following MAPI stores are supported: 
- Exchange mailboxes when using a online mode profile
- Outlook OST-files when using a cached mode profile
- PST files
- IMAP PST files
- Sharepoint PST stores
- Additional mailboxes, Automapped mailboxes
- Multi-Ex configuration (additional Exchange Accounts)
- Public folders

The tool is available as 32bit and 64bit version, the bitness need to match the installed Outlook version.
The tracing-information will be displayed in the tool and can be saved as a file for further analysis.


# **How to use this tool (simplified)**

1) Start the 'Outlook Monitor' tool and select the MAPI profile, which contains the MAPI stores you would like to monitor.
2) Expand the stores/folder hierarchy in the navigation pane and navigate to the store/folder, you would like to monitor
3) Double-click this folder to enable monitoring or Right-click this folder to enable monitoring or use the Commandbar-button in the toolbar to enable monitoring
(If registering this folder for notifications succeeded, the folder will be displayed as bold in the folder hierarchy. If registering failed, the folder won't be displayed as bold. In that case use the appropriate commandline-parameter for further troubleshooting).

Monitoring is now active and the "Outlook Monitor" will trace all changes.

4) After doing the repro, save the log file and close the tool. During close, all registered folders will be un-registered and all MAPI objects will be released.



# **Which information will be traced?**

The tool will trace and output all MAPI properties of an item, which has been added/changed/deleted.
Each change of a MAPI item will result in a new long entry for this change.
There are two modes to output item changes (Full Mode vs. Compare Mode - this can be controlled via Options-Use compare mode for Modified events).
See the 'Addition information' section for more details.



# **Scenarios: When to use the tool**

- Unexpected item processing in the inbox (Rules)
- External devices are syncing mailbox content and causing unwanted changes
- Unexpected calendar changes (duplicated calendar items, unexpected deletions…)
- To track and understand the workflows (e.g. what happens when sending an email)
- Exchange Backend services are modifying items (Calendar Assistant, Search Indexer, ....)
- Items disappear or get changed overnight/without any user connected to the mailbox (by some automated scripts….)
- Strange things happen (mail gets moved to deleted items, without touched by the user…)
- Custom Add-In/Script is changing items in the folder

The tool can only output the changes, it cannot tell which application/device/user has done the change (unless this information can be read from the MAPI properties of this item). Most Exchange backend-services can be identified by this.



# **How does the tracing work?**

This tool uses the "Event Notification in MAPI", see http://msdn.microsoft.com/en-us/library/office/cc842079.aspx for details. Folder events and content table events are supported.


|Folder event |Description |
|--|--|
|New mail (fnevNewMail) |A message has been delivered into a folder and is waiting to be processed|
|Object modified (fnevObjectModified)  |A MAPI object has changed  |
|Object created (fnevObjectCreated) |A MAPI object has been created |
|Object moved (fnevObjectMoved) |A MAPI object has been moved |
|Object deleted (fnevObjectDeleted) |A MAPI object has been deleted |
|Object copied (fnevObjectCopied) |A MAPI object has been copied|


|Table event|Description |
|--|--|
|TABLE_ROW_ADDED|A new row has been added to the content table and the corresponding object was saved |
|TABLE_ROW_MODIFIED|A row has been changed |
|TABLE_ROW_DELETED|A row has been removed from the table |

A Table event is an indicator for an item change, means the content in the Mapi folder has changed (the term "A row" means a single Mapi/Outlook item, e.g. mail item, calendar item, contact item.....)



# **The UI - the options**

- **Dump complete binary values**

This option is not enabled by default. The output of binary properties is limited to 1024 characters. If enabled, the complete value of the binary property will be written in the output.

- **Include Non-IPM subtree folders**

This option is not enabled by default, only sub-folders of IPM_SUBTREE will be added to the folder tree in the navigation pane. If enabled, all folders below the mailbox root will be added to the folder tree.
This option must be enabled BEFORE expanding a store item in the folder tree. Once the store has been expanded, checking/unchecking this option doesn't change the folder tree.

- **Use compare mode for MODIFIED events**

This option is enabled by default. In case of a TABLE_ROW_MODIFIED event, only changes to the previous state of the item will be tracked. Each single MAPI property of the changed item is compared with the previous state.
See the 'Additional information' section down below for further information.

- **Create full content copy on startup (slow)**

This option is not enabled by default. If enabled, the complete contents table of a MAPI folder will be read and kept in memory when registering a folder for notifications. If the folder contains many items, this is a long-running operation, which cannot be cancelled.
This option must be enabled BEFORE registering the first folder, it canned be turned on and off while monitoring is active.
See the 'Additional information' section down below for further information.

- **Dump items as msg for Add and Modify Events (slow)**

This option is not enabled by default. I enabled, the added/modified item will be stored as msg-file in the file system. A 'Logs' folder will be created in the same folder as the executable, which is used to store the msg-files.

- **Register associated content**

This option is not enabled by default. When registering a folder for table notifications, only the contents table will be registered. If this option is enabled, both contents table and the associated contents table will be registered.

- **Resolve named properties**

This option is enabled by default. Named properties will be displayed with the server-specific ID, and the global, server-independent ID.


- **Dump Recipients**

This option is not enabled by default. If enabled, all items in the recipient table for a given MAPI item with their properties will be logged.

- **Dump Attachments**

This option is not enabled by default. If enabled, all items in the attachment table for a given MAPI item with their properties will be logged.


# **Additional Information**

## The 'Compare mode' option

When the compare mode is ON, Outlook monitor compares the item properties BEFORE and AFTER the change, and only new properties (on this item), removed properties (on this item) and changed properties (on this item) will be listed. MAPI properties of the item, which haven't changed, won't be listed to simplify the output.

Comparing an item to its previous state can only work, if the 'Outlook Monitor' knows the previous state. One method is to enable the 'Create full content copy on startup' option, as this will read all existing items in the folder (and therefore 'Outlook Monitor' knows the previous state).
The other method (if this is sufficient for the troubleshooting) is to simply wait on the first change of the item. As soon as an item changes for the first time, Outlook keeps this item in memory and can compare its state for all subsequent changes.

Sample 1:
- A new item in the calendar is created (there is no previous state and therefore it is listed as new item with all properties)
- Something changes this item (As the previous state is known, the compare mode will work)

Sample 2 ('Create full content' is disabled):
- A change happens for an existing item (there is no previous state and therefore it is listed as new item with all properties)
- Something changes this item (As the previous state is known, the compare mode will work)

Sample 3 ('Create full content' is enabled):
- A change happens for an existing item (As the previous state is known because the content has been read while registering this folder, the compare mode will work and only changes to the previous state will be listed)

Sample 4 ('Create full content' is disabled):
- An existing item will be deleted. As the deleted item has been removed from the contents table, no information can be read and therefore nothing can be logged. If monitoring of the 'Deleted Items' folder is on, the deletion might be logged as a new item in this folder.

Sample 5 ('Create full content' is enabled):
- An existing item will be deleted. As the contents table has been read while registering the folder, the previous state of the item is known and therefore this state can be logged for the 'deleted' item.

## The 'Create full content' option

If you expect changes to existing items, or deletions of existing items, and if the folder doesn't contain too many items, this option should be enabled before monitoring this folder.
Be careful to monitor a whole store (an Exchange Mailbox, a Pubic Folder) when this 'Create full content' option is enabled. Due to the size of the store, reading all contents table will take long time, and cannot be cancelled.

## Command-line parameters

The following command-line parameter can be used:

### /EnableTS

-> this enables the troubleshooting view and extra output

### /Profile:<profile name>

-> the dialog to select a MAPI profile doesn't appear, the profile passed as <profile name> will be used

### /Default

-> the dialog to select a MAPI profile doesn't appear, the default MAPI profile will be used

(if both /Default and /Profile:<..> are set, /Default will win)


### /FldrList:<number>

-> this defines which folders will be monitored automatically after starting the Outlook Monitor tool

-> The default store is used for the selected MAPI profile


|Folder|Number|
|--|--|
|Inbox|0x1|
|Calendar|0x2|
|Contacts|0x4|
|Tasks|0x8|
|Draft|0x10|
|Outbox|0x20|
|Sent Items|0x40|
|Deleted Items|0x80|
|Complete Mailbox|0x100|

Summarize the numbers if you want to monitor more than one folder automatically, e.g. without letting the user to select the folders)

e.g. Inbox + Outbox, 0x1 + 0x20 = 0x21, decimal value: 33, Olmon3.exe /Fldrlist:33

e.g. Complete Mailbox, 0x100, decimal value: 256, Olmon3.exe /Fldrlist:256

Sample to start with the default profile and monitor Inbox + Outbox:
Olmon3.exe /Default /Fldrlist:33


## System requirements

MAPI sub system installed (will be installed with Outlook)


## Known issues, Tipps

As this tool relies on notifications, it’s recommended to add the executable to the windows firewall exception list. Otherwise you’ll lose notifications, or receive them delayed.

This tool does not require Outlook to run, but if you want to trace changes to the OST-file, Outlook has to run with this Cached Mode profile.

Several instances of the tool can run at the same time, on the same machine (e.g. Instance 1 with a Cached Mode profile, Instance2 with an Online Mode Profile - just to track how the sync to and from the mailbox works).

If an online mode profile is used, receiving notifications for an empty folder doesn't work. The workaround is to create one item in the folder.


## Bugs and Feedback

Feel free to ping me (thomno@microsoft.com).
