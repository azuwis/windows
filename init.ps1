$dir = "$PSScriptRoot"

function CreateShortcut {
    param($Shortcut, $TargetPath, $Arguments, $WindowStyle)
    if (-not (Test-Path $Shortcut)) {
        $ws = New-Object -ComObject ("WScript.Shell")
        $sc = $ws.CreateShortcut($Shortcut)
        $sc.TargetPath = $TargetPath
        $sc.WorkingDirectory = (Get-Item $TargetPath).Directory.FullName
        if (-not ($Arguments -eq $null)) { $sc.Arguments=$Arguments }
        if (-not ($WindowStyle -eq $null)) { $sc.WindowStyle=$WindowStyle }
        $sc.Save()
    }
}

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

function UnpackUrl {
    param($Url, $File, $UnpackDir, $TestDir)
    if (-not $File) {
        $File = $Url.Substring($Url.LastIndexOf("/") + 1)
        $Output = "$Home\Downloads\$File"
    }
    if (-not $TestDir) {
        $TestDir = $UnpackDir
    }
    if (-not (Test-Path "$TestDir")) {
        if (-not (Test-Path $Output)) {
            Import-Module BitsTransfer
            Start-BitsTransfer -Description "Downloading $File from $Url" -Source $Url -Destination $Output
        }
        switch ((Get-Item $Output).Extension) {
            '.zip' {
                $shell = new-object -com shell.application
                $shell.NameSpace($UnpackDir).CopyHere($shell.NameSpace($Output).Items())
            }
            '.7z' {
                 & "C:\Program Files\7-Zip\7z.exe" x "-o$UnpackDir" "$Output" | Out-Null
            }
        }
    }
}

# create programs dir
$programs = "C:\Programs"
if (-not (Test-Path $programs)) {
    mkdir -Path $programs
}

# disable bits branchcache https://powershell.org/forums/topic/bits-transfer-with-github/
if (-not (Get-ItemProperty HKLM:\SOFTWARE\Policies\Microsoft\Windows\BITS | Select-Object -ExpandProperty DisableBranchCache)) {
    Start-Process powershell -Verb runAs -ArgumentList  "New-ItemProperty HKLM:\SOFTWARE\Policies\Microsoft\Windows\BITS -Name DisableBranchCache -Value 1 -PropertyType DWORD -Force"
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

# install mpv
UnpackUrl -Url "https://cfhcable.dl.sourceforge.net/project/mpv-player-windows/64bit/mpv-x86_64-20180721-git-08a6827.7z" -UnpackDir "$programs\mpv"

# install weasel
InstallUrl -Name С�Ǻ�ݔ�뷨 -Url "https://dl.bintray.com/rime/weasel/weasel-0.11.1.0-installer.exe" -Arg /S

# enable wsl
if (-not (Test-Path "C:\Windows\System32\wsl.exe")) {
    Start-Process powershell -Verb runAs -ArgumentList "Enable-WindowsOptionalFeature -Online -FeatureName Microsoft-Windows-Subsystem-Linux"
}

# install debian
if (-not (Get-AppxPackage -Name TheDebianProject.DebianGNULinux)) {
    Start-Process "ms-windows-store://pdp/?ProductId=9MSVKQC78PK6"
}

# install qterminal
UnpackUrl -Url "https://github.com/kghost/qterminal/releases/download/0.9.0-wsl.1/QTerminal.X64.zip" -UnpackDir "$programs" -TestDir "$programs\QTerminal"

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
if (-not (Get-Process sshd -ErrorAction Ignore)) {
    Invoke-Item -Path $sshd
}

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
