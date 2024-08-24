# Exchange OU Attachment Size Modifier
# Author: Emad Mukhtar (https://github.com/emad-mukhtar)

# Function to test Exchange connection
function Test-ExchangeConnection {
    try {
        Get-ExchangeServer -ErrorAction Stop | Out-Null
        return $true
    }
    catch {
        return $false
    }
}

# Function to validate positive integer input
function Get-PositiveInteger {
    param ([string]$prompt)
    do {
        $value = Read-Host $prompt
        if ($value -match '^\d+$' -and [int]$value -gt 0) {
            return [int]$value
        }
        Write-Host "Please enter a positive integer." -ForegroundColor Yellow
    } while ($true)
}

# Display script information
Write-Host "Exchange OU Attachment Size Modifier" -ForegroundColor Cyan
Write-Host "Author: Emad Mukhtar (https://github.com/emad-mukhtar)" -ForegroundColor Cyan
Write-Host "This script modifies inbound & outbound email attachment size for all members of an Active Directory OU" -ForegroundColor Cyan

# Prompt for connection details
$server = Read-Host "Enter the name of the Exchange server"
$username = Read-Host "Enter the username for connecting to the Exchange server"
$password = Read-Host -AsSecureString "Enter the password for connecting to the Exchange server"
$adDomain = Read-Host "Enter the AD domain name"
$targetOU = Read-Host "Enter the target Active Directory OU name"

# Import module and connect to Exchange
try {
    Import-Module ExchangeOnlineManagement -ErrorAction Stop
    $cred = New-Object System.Management.Automation.PSCredential($username, $password)
    Connect-ExchangeOnline -Credential $cred -ErrorAction Stop
    Write-Host "Successfully connected to Exchange Online." -ForegroundColor Green
}
catch {
    Write-Host "Failed to connect to Exchange Online: $_" -ForegroundColor Red
    exit
}

# Verify Exchange connection
if (-not (Test-ExchangeConnection)) {
    Write-Host "Failed to establish a connection to Exchange." -ForegroundColor Red
    exit
}

# Prompt for new sizes
$newMaxReceiveSize = Get-PositiveInteger "Enter the new MaxReceiveSize in MB for all members of the target OU"
$newMaxSendSize = Get-PositiveInteger "Enter the new MaxSendSize in MB for all members of the target OU"

# Get all members of the target OU
try {
    $mailboxes = Get-Recipient -OrganizationalUnit $targetOU -ResultSize Unlimited -ErrorAction Stop
    Write-Host "Retrieved $($mailboxes.Count) mailboxes from the specified OU." -ForegroundColor Green
}
catch {
    Write-Host "Error retrieving mailboxes from the specified OU: $_" -ForegroundColor Red
    exit
}

# Process each mailbox
$successCount = 0
$failCount = 0

foreach ($mailbox in $mailboxes) {
    Write-Host "Processing mailbox: $($mailbox.Name)" -ForegroundColor Cyan
    
    try {
        $userQuotas = Get-Mailbox $mailbox.Name -ErrorAction Stop | 
                      Select-Object StorageLimitStatus,ProhibitSendQuota,IssueWarningQuota,ProhibitSendReceiveQuota
        
        if ($newMaxReceiveSize -gt $userQuotas.StorageLimitStatus.Value.ToMB() -or 
            $newMaxSendSize -gt $userQuotas.StorageLimitStatus.Value.ToMB()) {
            Write-Host "The specified max attachment size exceeds the mailbox size for $($mailbox.Name). Skipping." -ForegroundColor Yellow
            $failCount++
            continue
        }
        
        if ($newMaxReceiveSize -gt $userQuotas.ProhibitSendQuota.Value.ToMB() -or 
            $newMaxSendSize -gt $userQuotas.IssueWarningQuota.Value.ToMB() -or 
            $newMaxSendSize -gt $userQuotas.ProhibitSendReceiveQuota.Value.ToMB()) {
            Write-Host "The specified max attachment size exceeds quotas for $($mailbox.Name). Skipping." -ForegroundColor Yellow
            $failCount++
            continue
        }
        
        Set-Mailbox $mailbox.Name -MaxReceiveSize "$newMaxReceiveSize MB" -MaxSendSize "$newMaxSendSize MB" -ErrorAction Stop
        Write-Host "Successfully modified attachment sizes for $($mailbox.Name)" -ForegroundColor Green
        $successCount++
    }
    catch {
        Write-Host "Error processing mailbox $($mailbox.Name): $_" -ForegroundColor Red
        $failCount++
    }
}

# Display summary
Write-Host "`nOperation completed." -ForegroundColor Cyan
Write-Host "Successfully modified: $successCount mailboxes" -ForegroundColor Green
Write-Host "Failed to modify: $failCount mailboxes" -ForegroundColor Yellow

Write-Host "`nTo check the configured attachment size for a specific user, run:" -ForegroundColor Cyan
Write-Host "Get-Mailbox -Identity `"<UserName>`" | Select-Object MaxReceiveSize, MaxSendSize" -ForegroundColor Yellow

# Disconnect from Exchange
Disconnect-ExchangeOnline -Confirm:$false
Write-Host "Disconnected from Exchange Online." -ForegroundColor Green
