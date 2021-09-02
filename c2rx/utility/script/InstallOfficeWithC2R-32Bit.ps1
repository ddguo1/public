param (
    [string]
    $BuildNo
)

. "$PSScriptRoot\common\InstallOffice.ps1"

Write-Host("PSScriptRoot: $($PSScriptRoot)");

InstallOffice -UseMSIX $false -Platform x86 -Action Install -BuildNo $BuildNo -RootDir $PSScriptRoot