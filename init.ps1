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

function FirewallRule {
    param($DisplayName, $Action = "Allow", $Protocol = "TCP", $LocalPort)
    if (-not (Get-NetFirewallRule -DisplayName $DisplayName -ErrorAction Ignore)) {
        Start-Process powershell -Verb runAs -ArgumentList "New-NetFirewallRule -DisplayName $DisplayName -Action $Action -Protocol $Protocol -LocalPort $LocalPort"
    }
}

function InstallUrl {
    param($DisplayName, $Url, $Arg)
    $output = "$Home\Downloads\$DisplayName-installer.exe"
    if (-not (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*,HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* |  ? { $_.DisplayName -match $DisplayName })) {
        if (-not (Test-Path $output)) {
            Import-Module BitsTransfer
            Start-BitsTransfer -Description "Downloading $DisplayName installer from $Url" -Source $Url -Destination $Output
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
InstallUrl -DisplayName 7-Zip -Url "https://www.7-zip.org/a/7z1805-x64.exe" -Arg /S

# install firefox
InstallUrl -DisplayName Firefox -Url "https://download.mozilla.org/?product=firefox-latest&os=win64&lang=en-US" -Arg /S

# install mpv
UnpackUrl -Url "https://cfhcable.dl.sourceforge.net/project/mpv-player-windows/64bit/mpv-x86_64-20180721-git-08a6827.7z" -UnpackDir "$programs\mpv"

# install weasel
InstallUrl -DisplayName –°¿«∫¡›î»Î∑® -Url "https://dl.bintray.com/rime/weasel/weasel-0.11.1.0-installer.exe" -Arg /S

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
CreateShortcut -Shortcut ([Environment]::GetFolderPath("Desktop") + "\QTerminal.lnk") -TargetPath "$programs\QTerminal\QTerminal.exe"

# auto start sshd
CreateShortcut -Shortcut ([environment]::GetFolderPath("Startup") + "\sshd.lnk") -TargetPath C:\Windows\System32\wsl.exe -Arguments "sudo service ssh start" -WindowStyle 7
if (-not (Get-Process sshd -ErrorAction Ignore)) {
    Invoke-Item -Path $sshd
}

# allow sshd firewall inbound
FirewallRule -DisplayName WSL-OpenSSH-Server -LocalPort 22

# whitelist wsl in windows defender
Get-AppxPackage -Name TheDebianProject.DebianGNULinux | Select-Object -ExpandProperty PackageFamilyName | % {
    $path = "$env:LOCALAPPDATA\Packages\$_"
    if (-not (Get-MpPreference | Select-Object -ExpandProperty ExclusionPath) -contains $path) {
        Start-Process powershell -Verb runAs -ArgumentList "Set-MpPreference -ExclusionPath $path"
    }
}
