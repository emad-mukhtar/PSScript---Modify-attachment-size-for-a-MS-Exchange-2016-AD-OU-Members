# PSScript---Modify-attachment-size-for-a-MS-Exchange-2016-AD-OU-Members

# PowerShell script to modify MaxReceiveSize & MaxSendSize for all members of an AD OU
## Introduction
This PowerShell script allows you to modify the MaxReceiveSize and MaxSendSize properties of all members of an Active Directory Organizational Unit (OU).

## Requirements
Microsoft Exchange Server 2016 or later
Exchange Management Shell
Active Directory account with the necessary permissions to modify the MaxReceiveSize and MaxSendSize properties of Exchange mailboxes
The Active Directory module for Windows PowerShell
## How to Use
Open the Exchange Management Shell
Run the script by typing .\script-name.ps1
Enter the name of the Active Directory domain
Enter the name of the Exchange server
Enter your Active Directory username and password for connecting to the Exchange server
Enter the new MaxReceiveSize and MaxSendSize in MB for all members of the target OU
The script will then display a confirmation message and ask for confirmation to proceed with the changes. Type "OK" to confirm or "CANCEL" to cancel the operation.
## Note
The script will modify the MaxReceiveSize and MaxSendSize properties of all mailboxes in the target OU. Ensure that you have selected the correct OU before proceeding with the operation.
The script adds fail-safe checks to ensure that the specified MaxReceiveSize and MaxSendSize values do not exceed the maximum storage size or quota limits of the user's mailbox.
## Contact
If you have any questions or feedback, please visit my GitHub at https://github.com/emad-mukhtar.

Please use this script at your own risk. I'm not responsible for any damage or loss caused by the use of this script. It is highly recommended to backup your Exchange data before running this script.
