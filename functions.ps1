###Includes###
. .\statics.ps1

function InitializeModules() {
    Import-Module ExchangeOnlineManagement
    Connect-ExchangeOnline -UserPrincipalName $adminUpn
    Connect-MsolService
}


function RemoveAccess() {
    param(
        [Parameter(Mandatory = $true)] [string] $upn
    )

    RemoveMeetLicense($upn)
    RemoveApplicationLicenses($upn)
    DisableExchangeActiveSync($upn)
    DisableOwaDevices($upn)
    DisableOwaAccess($upn)
}

function RemoveMeetLicense(){
param(
    [Parameter(Mandatory = $true)] [string] $upn
)
    try {
        Set-MsolUserLicense -UserPrincipalName $upn -RemoveLicenses supplychainguru:MCOMEETADV -ErrorAction "Stop" 
    
        $logString = "Removed the meet license for " + $upn 
        WriteToLogFile ([LogType]::DebugLog) $logString
    
    }
    catch {
        $logString = "Failed to remove the meet license for  " + $upn + ", they probably didn't have it"
        WriteToLogFile ([LogType]::ExceptionsLog) $logString
    }
}

function RemoveApplicationLicenses() {
    param(
        [Parameter(Mandatory = $true)] [string] $upn
    )
    try {
        Set-MsolUserLicense -UserPrincipalName $upn -LicenseOptions $LicensesWithoutTeams -ErrorAction "Stop" 
        $logString = "Changed the licenses for " + $upn 
        WriteToLogFile ([LogType]::DebugLog) $logString
    
    }
    catch {
        $logString = "Failed to apply license settings for " + $upn
        WriteToLogFile ([LogType]::ExceptionsLog) $logString
    }
}

function DisableExchangeActiveSync() {
    param(
        [Parameter(Mandatory = $true)] [string] $upn
    )
    try {
        Set-CasMailbox -Identity $upn -ActiveSyncEnabled $false -ErrorAction "Stop" 
        $logString = "Changed Exchange Activesync for " + $upn 
        WriteToLogFile ([LogType]::DebugLog) $logString
    }
    catch {
        $logString = "Failed to disable exchange activesync for  " + $upn
        WriteToLogFile ([LogType]::ExceptionsLog) $logString
    }
}

function DisableOwaDevices() {
    param(
        [Parameter(Mandatory = $true)] [string] $upn
    )
    try {
        Set-CasMailbox -Identity $upn -OWAforDevicesEnabled $false -ErrorAction "Stop" 
        $logString = "Changed OWA device access for  " + $upn 
        WriteToLogFile ([LogType]::DebugLog) $logString
    }
    catch {
        $logString = "Failed to disable OWA device settings for  " + $upn
        WriteToLogFile ([LogType]::ExceptionsLog) $logString
    }
}

function DisableOwaAccess() {
    param(
        [Parameter(Mandatory = $true)] [string] $upn
    )

    try {
        Set-CasMailbox -Identity $upn -OWAEnabled $false -ErrorAction "Stop" 
        $logString = "Changed OWA license for  " + $upn 
        WriteToLogFile ([LogType]::DebugLog) $logString
    }
    catch {
        $logString = "Failed to disable the OWA license for " + $upn
        WriteToLogFile ([LogType]::ExceptionsLog) $logString
    }
}

enum LogType {
    DebugLog
    ExceptionsLog
}
function WriteToLogFile() {
    param
    (
        [Parameter(Mandatory = $true)] [LogType] $typeOfLog,
        [Parameter(Mandatory = $true)] [string] $stringToWrite
    )

    switch ($typeOfLog) {
        ([LogType]::DebugLog) { 
            $destLocation = $debugLog
            Add-Content -Path .\$destLocation -Value  $stringToWrite
        }
        ([LogType]::ExceptionsLog) {
            $destLocation = $exceptionsLog
            Add-Content -Path .\$destLocation -Value  $stringToWrite
        }
        Default {
            Write-Host "You didn't input a correct type of log, make sure you use a LOGTYPE"
        }
    }
}