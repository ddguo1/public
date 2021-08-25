@reg add "HKEY_LOCAL_MACHINE\Software\Microsoft\ClickToRun\OverRide" /v UseMSIXVirtualization /t REG_DWORD /d 0 /f

@"\\yu-pc1.redmond.corp.microsoft.com\Shared\msix\addinValidation\setup.exe" /configure "\\yu-pc1.redmond.corp.microsoft.com\Shared\msix\addinValidation\removeAll.xml"