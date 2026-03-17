# OneDrive Site Collection and Administrator Management Script
# Purpose: Manage site collections and administrators for OneDrive instances
# Author: Script Assistant
# Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')

param(
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
            if ($userPart -match "^(.+)_([^_]+)_([^_]+)$") {
                $localPart = $Matches[1] -replace "_", "."
                $domainName = $Matches[2]
                $domainExt = $Matches[3]
                $UserUPN = "$localPart@$domainName.$domainExt"
                
                Write-Log "Extracted UPN: $UserUPN" "INFO"
                
                # Remove Filelock from site collection admin
                Write-Log "Removing $FileLockAccount from site collection admins" "INFO"
                Set-SPOUser -Site $Site.Url -LoginName $FileLockAccount `
                    -IsSiteCollectionAdmin $false -ErrorAction SilentlyContinue
                Write-Log "Successfully removed $FileLockAccount from admins" "SUCCESS"
                
                # Set original owner as site collection admin
                Write-Log "Setting $UserUPN as site collection admin" "INFO"
                Set-SPOUser -Site $Site.Url -LoginName $UserUPN `
                    -IsSiteCollectionAdmin $true -ErrorAction Stop
                Write-Log "Successfully set $UserUPN as admin" "SUCCESS"
            }
            else {
                Write-Log "Failed to parse URL format for: $userPart" "WARNING"
            }
        }
        catch {
            Write-Log "Error processing $($Site.Url): $($_.Exception.Message)" "ERROR"
        }
    }

    Write-Progress -Activity "Processing OneDrive sites" -Completed

# ...existing code...