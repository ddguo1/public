param (
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [Boolean]$UseMSIX,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [ValidateSet('Install', 'Upsell', 'Update')]
        [string]$Action)

. "$PSScriptRoot\common\utility.ps1"

write-host "InstallOffice.ps1: UseMSIX: $($UseMSIX); Action: $($Action)";
Write-Host("PSScriptRoot: $($PSScriptRoot)");

EnableMSIXVirtualization -Enable $UseMSIX -ForAction $Action
