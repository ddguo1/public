function EnableMSIXVirtualization {
    param (
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [Boolean]$Enable,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [ValidateSet('Install', 'Upsell', 'Downsell', 'Update')]
        [string]$ForAction
    )

    if ($ForAction -eq "Update") {
        New-Item -Path "HKLM:\Software\Microsoft\ClickToRun\OverRide" -force
        New-ItemProperty -Path "HKLM:\Software\Microsoft\ClickToRun\OverRide" -Name "UseMSIXVirtualizationForUpdate" -PropertyType "DWord" -Value $Enable -force;
        return;
    }

    if ($ForAction -eq "Install" -or $ForAction -eq "Upsell" -or $ForAction -eq "Downsell") {
        New-Item -Path "HKLM:\Software\Microsoft\ClickToRun\OverRide" -force
        New-ItemProperty -Path "HKLM:\Software\Microsoft\ClickToRun\OverRide" -Name "UseMSIXVirtualization" -PropertyType "DWord" -Value $Enable -force;
        return;
    }

    #
    # OptOutMsixVirtualization
}

function EnableOfficeInsiders{
    param (
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [Boolean]$Enable
    )

    New-Item -Path "HKCU:\SOFTWARE\Policies\Microsoft\office\16.0\Common" -force
    New-ItemProperty -Path "HKCU:\SOFTWARE\Policies\Microsoft\office\16.0\Common" -Name "EnableExtendedInsidersList" -PropertyType "DWord" -Value $Enable -force
    New-ItemProperty -Path "HKCU:\SOFTWARE\Policies\Microsoft\office\16.0\Common" -Name "DisableInsiderSlabForInternalUsers" -PropertyType "DWord" -Value $Enable -force
    if ($Enable) {
        New-ItemProperty -Path "HKCU:\SOFTWARE\Policies\Microsoft\office\16.0\Common" -Name "insiderslabbehavior" -PropertyType "DWord" -Value 1 -force
    }
    # New-ItemProperty insiderslabbehavior    2 // exist even office has not been installed yet.
}
