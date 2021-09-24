function InstallOffice {
    param (
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [Boolean]$UseMSIX,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [ValidateSet('x64', 'x86')] # , 'arm64'
        [string]$Platform,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [ValidateSet('Install', 'Upsell', 'Update')]
        [string]$Action,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]
        $RootDir,

        [string]
        $BuildNo,

        [string]
        $ConfigXmlStencil = $null
    )

    # Install/Upsell/Downsell
    #    App-V/MSIX
    #    x86/x64/arm64
    #    version
    # Uninstall

    . "$PSScriptRoot\utility.ps1"

    write-host Get-Location

    write-host "InstallOffice.ps1: UseMSIX: $($UseMSIX); Platform: $($Platform); BuildNo: $($BuildNo)";
    Write-Host("RootDir: $($RootDir)");


    if ($Action -eq 'Update') {
        Write-host "Office couldn't be updated with this powershell script!" -ForegroundColor Red
        Write-host "Please first select between MSIX and App-V by script EnableMSIXForOffice.ps1" -ForegroundColor Red
        Write-host "Then Update Office in any Office App: File->Account->Update Options->Update Now!" -ForegroundColor Red

        return;
    }

    if ($BuildNo -and (-not ($BuildNo -match '16\.0\.1\d{4}\.\d{5}'))) {
        Write-host "BuildNo ($(BuildNo)) is invalid." -ForegroundColor Red
        return;
    }

    EnableOfficeInsiders -Enable $true
    EnableMSIXVirtualization -Enable $UseMSIX -ForAction $Action

    $executablePath = Join-Path -Path $RootDir -ChildPath "setup.exe";
    $configXmlStencilPath = Join-Path -Path $RootDir -ChildPath "config\Office.xml";

    if ($ConfigXmlStencil) {
        $configXmlStencilPath = Join-Path -Path $RootDir -ChildPath "config\$($ConfigXmlStencil)";
    }
    # write a concrete config.xml in the temporary folder.
    $configXmlPath = Join-Path -Path $env:temp -ChildPath "Office-$(get-date -f yyyy.MM.dd.HH.mm.ss.ffff).xml"

    write-host "configxmlPath: $($configXmlPath)"

    $OfficeClientEditions = @{
        x64="64"
        x86="32"
    }

    $Channel = "BetaChannel"
    $OfficeClientEdition=$OfficeClientEditions[$Platform]
    $Description = "Office $($OfficeClientEdition)-bit $($Channel) Setup"

    $versionReplacement = "";
    if ($BuildNo) {
        $versionReplacement = "Version='$($BuildNo)'";
    }

    $content = [System.IO.File]::ReadAllText($configXmlStencilPath)
    $content = $content.Replace("%Description%", $Description).
                        Replace("%OfficeClientEdition%", $OfficeClientEdition).
                        Replace("%Channel%", $Channel).
                        Replace("%Version%", $versionReplacement);

    [System.IO.File]::WriteAllText($configXmlPath, $content);

    write-host $content

    $p = Start-Process -FilePath $executablePath -ArgumentList "/configure",$configXmlPath -wait -PassThru #-NoNewWindow
    $p.HasExited
    $p.ExitCode
}

function UnInstallOffice {
    param (
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]
        $RootDir
    )
    $executablePath = Join-Path -Path $RootDir -ChildPath "setup.exe";
    $configXmlPath = Join-Path -Path $RootDir -ChildPath "config\removeAll.xml";
    Start-Process -FilePath $executablePath -ArgumentList "/configure",$configXmlPath -wait -PassThru #-NoNewWindow

    EnableMSIXVirtualization -Enable $false -ForAction Install
}
