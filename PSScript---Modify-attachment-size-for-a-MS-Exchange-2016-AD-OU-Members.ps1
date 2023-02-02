Write-Host "Thank you for using my script, visit my GitHub for more https://github.com/emad-mukhtar"
Write-Host "This is a PowerShell script to modify inbound & outbound email attachment size for all members of an active directory OU"

# Prompt for server name, username, password, AD domain name, and the target OU name
$server = Read-Host "Enter the name of the Exchange server"
$username = Read-Host "Enter the username for connecting to the Exchange server"
$password = Read-Host -AsSecureString "Enter the password for connecting to the Exchange server"
$adDomain = Read-Host "Enter the AD domain name"
$targetOU = Read-Host "Enter the target active directory OU name"

# Import the Exchange PowerShell module and connect to the Exchange server
Import-Module ExchangeOnlineManagement
$cred = New-Object System.Management.Automation.PSCredential($username, $password)
Connect-ExchangeOnline -Credential $cred

# Prompt for the new MaxReceiveSize and MaxSendSize in MB
$newMaxReceiveSize = Read-Host "Enter the new MaxReceiveSize in MB for all members of the target OU"
$newMaxSendSize = Read-Host "Enter the new MaxSendSize in MB for all members of the target OU"

# Get all members of the target OU
$mailboxes = Get-Recipient -OrganizationalUnit $targetOU

# Loop through all members of the target OU
foreach ($mailbox in $mailboxes) {
    # Adding fail safe checks
    $userQuotas = Get-Mailbox $mailbox.Name | Select-Object StorageLimitStatus,ProhibitSendQuota,IssueWarningQuota,ProhibitSendReceiveQuota
    if($newMaxReceiveSize -gt $userQuotas.StorageLimitStatus.Value.ToMB() -or $newMaxSendSize -gt $userQuotas.StorageLimitStatus.Value.ToMB()){
        Write-Host "The specified max attachment size exceeds the user's mailbox size for $($mailbox.Name). Please choose a smaller size."
        continue
    }
    if($newMaxReceiveSize -gt $userQuotas.ProhibitSendQuota.Value.ToMB() -or $newMaxSendSize -gt $userQuotas.IssueWarningQuota.Value.ToMB() -or $newMaxSendSize -gt $userQuotas.ProhibitSendReceiveQuota.Value.ToMB()){
        Write-Host "The specified max attachment size exceeds the user's ProhibitSend, IssueWarning, or ProhibitSendReceive quotas for $($mailbox.Name). Please choose a smaller size."
        continue
    }
    
    # Modify the MaxReceiveSize and MaxSendSize for the current user
    Set-Mailbox $mailbox.Name -MaxReceiveSize $newMaxReceiveSize'MB' -MaxSendSize $newMaxSendSize'MB'
    Write-Host "MaxReceiveSize and MaxSendSize for $($mailbox.Name) have been successfully modified to $newMaxReceiveSize MB and $newMaxSendSize MB, respectively"
}

Write-Host "You can always run the following command to check the currently configured Attachment Size: Get-Mailbox -Identity "<UserName>" | Select MaxReceiveSize, MaxSendSize"
