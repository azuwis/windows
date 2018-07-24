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
