$dir = "$PSScriptRoot"

# swap capslock ctrl
try {
    Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Keyboard Layout" | Select-Object -ExpandProperty "Scancode Map" -ErrorAction Stop | Out-Null
} catch {
    regedit /s "$dir\switch-capslock-ctrl.reg"
}

# enable wsl
if (-not (Test-Path "C:\Windows\System32\wsl.exe")) {
    Start-Process powershell -Verb runAs -ArgumentList "Enable-WindowsOptionalFeature -Online -FeatureName Microsoft-Windows-Subsystem-Linux"
}

# install debian
if (-not (Get-AppxPackage -Name TheDebianProject.DebianGNULinux)) {
    Start-Process "ms-windows-store://pdp/?ProductId=9MSVKQC78PK6"
}

# install qterminal
$url = "https://github.com/kghost/qterminal/releases/download/0.9.0-wsl.1/QTerminal.X64.zip"
$zip = "$Home\Downloads\QTerminal.X64.zip"
$dir = "C:\Programs"
if (-not (Test-Path "$dir\QTerminal")) {
    if (-not (Test-Path $zip)) {
        Import-Module BitsTransfer
        Start-BitsTransfer -Source $url -Destination $zip
    }
    $shell = new-object -com shell.application
    $shell.NameSpace($dir).copyhere($shell.NameSpace($zip).Items())
}

# auto start sshd
$sshd = [environment]::getfolderpath("Startup") + "\sshd.lnk"
if (-Not (Test-Path $sshd)) {
    $wscript = New-Object -ComObject ("WScript.Shell")
    $shortcut = $wscript.CreateShortcut($sshd)
    $shortcut.TargetPath="C:\Windows\System32\wsl.exe"
    $shortcut.Arguments="sudo service ssh start"
    $shortcut.WorkingDirectory = "C:\Windows\System32"
    $shortcut.WindowStyle = 7
    $shortcut.Save()
}
Invoke-Item -Path $sshd
