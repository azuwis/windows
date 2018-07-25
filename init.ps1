$dir = "$PSScriptRoot"

# swap capslock ctrl
try {
    Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Keyboard Layout" | Select-Object -ExpandProperty "Scancode Map" -ErrorAction Stop | Out-Null
} catch {
    regedit /s "$dir\switch-capslock-ctrl.reg"
}

# install firefox
$url = "https://download.mozilla.org/?product=firefox-latest&os=win64&lang=en-US"
$output = "$Home\Downloads\firefox-installer.exe"
if (-not (Test-Path "C:\Program Files\Mozilla Firefox")) {
    if (-not (Test-Path $output)) {
        Import-Module BitsTransfer
        Start-BitsTransfer -Source $url -Destination $output
    }
    Start-Process $output -ArgumentList /S
}

# install weasel
$url = "https://dl.bintray.com/rime/weasel/weasel-0.11.1.0-installer.exe"
$output = "$Home\Downloads\weasel-0.11.1.0-installer.exe"
if (-not (Test-Path "C:\Program Files (x86)\Rime")) {
    if (-not (Test-Path $output)) {
        Import-Module BitsTransfer
        Start-BitsTransfer -Source $url -Destination $output
    }
    Start-Process $output -ArgumentList /S
}

# create programs dir
$programs = "C:\Programs"
if (-not (Test-Path $programs)) {
    mkdir -Path $programs
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
if (-not (Test-Path "$programs\QTerminal")) {
    if (-not (Test-Path $zip)) {
        Import-Module BitsTransfer
        Start-BitsTransfer -Source $url -Destination $zip
    }
    $shell = new-object -com shell.application
    $shell.NameSpace($programs).copyhere($shell.NameSpace($zip).Items())
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

# allow sshd firewall inbound
if (-not (Get-NetFirewallRule -DisplayName "WSL-OpenSSH-Server" -ErrorAction Ignore)) {
    Start-Process powershell -Verb runAs -ArgumentList "New-NetFirewallRule -DisplayName WSL-OpenSSH-Server -Protocol TCP -LocalPort 22 -Action Allow"
}
