param (
    [string]
    $BuildNo
)

. "$PSScriptRoot\common\InstallOffice.ps1"

Write-Host("PSScriptRoot: $($PSScriptRoot)");

InstallOffice -UseMSIX $true -Platform x64 -Action Install -BuildNo $BuildNo -RootDir $PSScriptRoot -ConfigXmlStencil Office.with.Visio.xml