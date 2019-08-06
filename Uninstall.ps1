# Uninstalls an addon through VSTO

$installerPath = Join-Path $env:CommonProgramFiles 'microsoft shared\VSTO\10.0\VSTOInstaller.exe'
$parameter = @('/uninstall',  'file:///Sidekick.vsto')
& $installerPath @parameter

# Remove file that specifies which version the program is
$version = "3.0.1.221"
Remove-Item -Path "$env:ProgramData\Sidekick\$version.txt" -Force
