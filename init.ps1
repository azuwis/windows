$dir = "$PSScriptRoot"

function InstallUrl {
    param($Name, $Url, $Arg)
    $output = "$Home\Downloads\$Name-installer.exe"
    if (-not (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*,HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* |  ? { $_.DisplayName -match $Name })) {
        if (-not (Test-Path $output)) {
            Import-Module BitsTransfer
            Start-BitsTransfer -Description "Downloading $Name installer from $Url" -Source $Url -Destination $Output
        }
        if ($Arg) {
            Start-Process $output -ArgumentList $Arg
        } else {
            Start-Process $output
        }
    }
}

# create programs dir
$programs = "C:\Programs"
if (-not (Test-Path $programs)) {
    mkdir -Path $programs
}

# swap capslock ctrl
try {
    Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Control\Keyboard Layout" | Select-Object -ExpandProperty "Scancode Map" -ErrorAction Stop | Out-Null
} catch {
    regedit /s "$dir\switch-capslock-ctrl.reg"
}

# install 7z
InstallUrl -Name 7-Zip -Url "https://www.7-zip.org/a/7z1805-x64.exe" -Arg /S

# install firefox
InstallUrl -Name Firefox -Url "https://download.mozilla.org/?product=firefox-latest&os=win64&lang=en-US" -Arg /S

# install weasel
InstallUrl -Name –°¿«∫¡›î»Î∑® -Url "https://dl.bintray.com/rime/weasel/weasel-0.11.1.0-installer.exe" -Arg /S

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

# whitelist wsl in windows defender
Get-AppxPackage -Name TheDebianProject.DebianGNULinux | Select-Object -ExpandProperty PackageFamilyName | % {
    $path = "$env:LOCALAPPDATA\Packages\$_"
    if (-not (Get-MpPreference | Select-Object -ExpandProperty ExclusionPath) -contains $path) {
        Start-Process powershell -Verb runAs -ArgumentList "Set-MpPreference -ExclusionPath $path"
    }
}
