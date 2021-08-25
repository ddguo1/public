@reg add "HKEY_LOCAL_MACHINE\Software\Microsoft\ClickToRun\OverRide" /v UseMSIXVirtualization /t REG_DWORD /d 1 /f
@reg add "HKCU\SOFTWARE\Policies\Microsoft\office\16.0\common" /v EnableExtendedInsidersList /d 1 /t REG_DWORD /f /reg:64

@"\\yu-pc1.redmond.corp.microsoft.com\Shared\msix\addinValidation\setup.exe" /configure "\\yu-pc1.redmond.corp.microsoft.com\Shared\msix\addinValidation\Office-32Bit-BetaChannel.xml"
