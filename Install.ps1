<#
    .DESCRIPTION: THIS PROGRAM IS GARBAGE, WRAP IT IN HUBSPOT
                  Script to install Hubspot.
    .AUTHOR: Fredrik Standahl @ Syscom AS
    .Date: 06.08.2019 

#>

# Is HubSpot Sidekick installed? Look through Office Addins, impossible to find it otherwise

$searchScopes = "HKCU:\SOFTWARE\Microsoft\Office\Outlook\Addins","HKLM:\SOFTWARE\Wow6432Node\Microsoft\Office\Outlook\Addins"
$names = $searchScopes | % {Get-ChildItem -Path $_ | % {Get-ItemProperty -Path $_.PSPath} | Select-Object @{n="Name";e={Split-Path $_.PSPath -leaf}},FriendlyName,Description} | Sort-Object -Unique -Property name

if("Sidekick" -in $names.name){
    # Uninstalls Sidekick addon through VSTO

    $installerPath = Join-Path $env:CommonProgramFiles 'microsoft shared\VSTO\10.0\VSTOInstaller.exe'
    $parameter = @('/uninstall',  "file:///$PSScriptRoot\Sidekick_previous_update.vsto")
    & $installerPath @parameter

    # Remove file that contains version number

    $version = "3.0.1.221"
    Remove-Item -Path "$env:ProgramData\Sidekick\$version.txt" -Force

 }

# Installs Sidekick addon through VSTO

$installerPath = Join-Path $env:CommonProgramFiles 'microsoft shared\VSTO\10.0\VSTOInstaller.exe'
$parameter = @('/install',  "file:///$PSScriptRoot\Sidekick.vsto")
& $installerPath @parameter

# Creates a file that specifies which version the program is. Impossible to find out otherwise. 
# This is used for SCCM detection
$version = "3.0.1.221"
New-Item -Path "$env:ProgramData\Sidekick" -ItemType File -Name "$version.txt" -Force