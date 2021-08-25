param (
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [Boolean]$UseMSIX,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [ValidateSet('x64', 'x86')] # , 'arm64'
        [string]$Processor,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [ValidateSet('Install', 'Upsell', 'Update')]
        [string]$Action,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]
        $BuildNo)

# Install/Upsell/Downsell
#    App-V/MSIX
#    x86/x64/arm64
#    version
# Uninstall

. "$PSScriptRoot\common\utility.ps1"

write-host "InstallOffice.ps1: UseMSIX: $($UseMSIX); Processor: $($Processor); BuildNo: $($BuildNo)";
Write-Host("PSScriptRoot: $($PSScriptRoot)");


if ($Action -eq 'Update') {
    Write-host "Office couldn't be updated with this powershell script!" -ForegroundColor Red
    Write-host "Please first select between MSIX and App-V by script EnableMSIXForOffice.ps1" -ForegroundColor Red
    Write-host "Then Update Office in any Office App: File->Account->Update Options->Update Now!" -ForegroundColor Red

    return;
}

EnableOfficeInsiders -Enable $true
EnableMSIXVirtualization -Enable $UseMSIX -ForAction $Action

# $executablePath = Join-Path -Path $PSScriptRoot -ChildPath "setup.exe";
# $configXmlPath = Join-Path -Path $PSScriptRoot -ChildPath "Office-32Bit-BetaChannel.xml";

# $p = Start-Process -FilePath $executablePath -ArgumentList "/configure",$configXmlPath -wait -PassThru #-NoNewWindow
# $p.HasExited
# $p.ExitCode
