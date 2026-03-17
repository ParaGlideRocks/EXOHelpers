# OneDrive Site Collection and Administrator Management Script
# Purpose: Manage site collections and administrators for OneDrive instances
# Author: Script Assistant
# Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')

param(
    [Parameter(Mandatory = $true)]
    [string]$AdminUrl,

    [Parameter(Mandatory = $true)]
    [string]$FileLockAccount,

    [string]$LogPath = "$PSScriptRoot\OneDrive_AdminManagement_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
)

# Initialize logging function
# ...existing code...

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('INFO', 'WARNING', 'ERROR', 'SUCCESS')]
        [string]$Level = 'INFO'
    )
    
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $logMessage = "[$timestamp] [$Level] $Message"

    $color = switch ($Level) {
        'INFO'    { 'Cyan' }
        'WARNING' { 'Yellow' }
        'ERROR'   { 'Red' }
        'SUCCESS' { 'Green' }
        default   { 'White' }
    }
    
    Write-Host $logMessage -ForegroundColor $color
    Add-Content -Path $LogPath -Value $logMessage -ErrorAction SilentlyContinue
}

# ...existing code...

# URL Admin del tenant
# Connessione
Connect-SPOService -Url $AdminUrl

# Recupera tutti i OneDrive
$OneDrives = Get-SPOSite -IncludePersonalSite $true -Limit All -Filter "Url -like '-my.sharepoint.com/personal/'" | ? { $_.Owner -eq $FileLockAccount }

# Process each OneDrive
$totalSites = $OneDrives.Count
$currentIndex = 0

foreach ($Site in $OneDrives) {
    $currentIndex++
    $percentComplete = [int](($currentIndex / $totalSites) * 100)

    Write-Progress `
        -Activity "Processing OneDrive sites" `
        -Status "Processing $currentIndex of $totalSites : $($Site.Url)" `
        -PercentComplete $percentComplete

    try {
        Write-Log "Processing OneDrive: $($Site.Url)" "INFO"
        
        $url = $Site.Url
        $userPart = ($Site.Url -split "/")[-1]
        
        # Parse user information from URL
        if ($userPart -match "^(.+?)_([a-zA-Z0-9]+)_([a-zA-Z0-9]+)_([a-zA-Z]{2,})$") {
                # Subdomain format: user_tenant_onmicrosoft_com
                $localPart  = $Matches[1] -replace "_", "."
                $domainName = "$($Matches[2]).$($Matches[3])"
                $domainExt  = $Matches[4]
                $UserUPN = "$localPart@$domainName.$domainExt"
            }
            elseif ($userPart -match "^(.+)_([^_]+)_([^_]+)$") {
                # Standard format: user_domain_com
                $localPart  = $Matches[1] -replace "_", "."
                $domainName = $Matches[2]
                $domainExt  = $Matches[3]
                $UserUPN = "$localPart@$domainName.$domainExt"
            }
            else {
                Write-Log "ERROR: Could not parse URL format for: $userPart" "ERROR"
                Exit
            }
            
            Write-Log "Extracted UPN: $UserUPN" "INFO"
            
            # Set original owner as site collection admin
            Write-Log "Setting $UserUPN as site collection admin" "INFO"
            $setSPOUserParams = @{
                Site                  = $Site.Url
                LoginName             = $UserUPN
                IsSiteCollectionAdmin = $true
                ErrorAction           = 'Stop'
            }
            Write-log "Set-SPOUser @setSPOUserParams" "INFO"

            Write-Log "Successfully set $UserUPN as admin" "SUCCESS"

            # Remove Filelock from site collection admin
            Write-Log "Removing $FileLockAccount from site collection admins" "INFO"
            $removeFileLockParams = @{
                Site                  = $Site.Url
                LoginName             = $FileLockAccount
                IsSiteCollectionAdmin = $false
                ErrorAction           = 'SilentlyContinue'
            }
            Write-log "Set-SPOUser @removeFileLockParams" "INFO"
            Write-Log "Successfully removed $FileLockAccount from admins" "SUCCESS"
            
    }
    catch {
        Write-Log "Error processing $($Site.Url): $($_.Exception.Message)" "ERROR"
    }
}

Write-Progress -Activity "Processing OneDrive sites" -Completed

# ...existing code...