. "$PSScriptRoot\common\utility.ps1"

Write-Host("PSScriptRoot: $($PSScriptRoot)");

EnableOfficeInsiders -Enable $true
EnableMSIXVirtualization -Enable $false -ForAction 'Install'

$executablePath = Join-Path -Path $PSScriptRoot -ChildPath "setup.exe";
$configXmlPath = Join-Path -Path $PSScriptRoot -ChildPath "Office-32Bit-BetaChannel.xml";

$p = Start-Process -FilePath $executablePath -ArgumentList "/configure",$configXmlPath -wait -PassThru #-NoNewWindow
$p.HasExited
$p.ExitCode
