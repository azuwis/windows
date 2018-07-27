# create programs dir
$Desktop = [Environment]::GetFolderPath("Desktop")
$Startup = [Environment]::GetFolderPath("Startup")
$Programs = "$Env:HOMEDRIVE:\Programs"
if (-not (Test-Path $Programs)) {
    mkdir -Path $Programs
}

function CreateShortcut {
    param($Shortcut,$TargetPath,$Arguments,$WindowStyle)
    if (-not (Test-Path $Shortcut)) {
        $ws = New-Object -ComObject ("WScript.Shell")
        $sc = $ws.CreateShortcut($Shortcut)
        $sc.TargetPath = $TargetPath
        $sc.WorkingDirectory = (Get-Item $TargetPath).Directory.FullName
        if (-not ($Arguments -eq $null)) { $sc.Arguments = $Arguments }
        if (-not ($WindowStyle -eq $null)) { $sc.WindowStyle = $WindowStyle }
        $sc.Save()
    }
}

function DefenderExcludeAppx {
    param($Name)
    Get-AppxPackage -Name $Name | Select-Object -ExpandProperty PackageFamilyName | ForEach-Object {
        DefenderExcludePath -Path "$Env:LOCALAPPDATA\Packages\$_"
    }
}

function DefenderExcludePath {
    param($Path)
    if (-not (Get-MpPreference | Select-Object -ExpandProperty ExclusionPath) -contains $Path) {
        RunAsAdmin "Set-MpPreference -ExclusionPath `"$Path`""
    }
}

function FirewallRule {
    param($DisplayName,$Action = "Allow",$Protocol = "TCP",$LocalPort)
    if (-not (Get-NetFirewallRule -DisplayName $DisplayName -ErrorAction Ignore)) {
        RunAsAdmin "New-NetFirewallRule -DisplayName `"$DisplayName`" -Action `"$Action`" -Protocol `"$Protocol`" -LocalPort `"$LocalPort`""
    }
}

function InstallUrl {
    param($DisplayName,$Url,$Arg)
    $output = "$Home\Downloads\$DisplayName-installer.exe"
    if (-not (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*,HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object { $_.DisplayName -match $DisplayName })) {
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

function Registry {
    param($Path,$Name,$Value,$Type)
    $Json = (ConvertTo-Json $Value)
    if (-not (Test-Path $Path)) {
        RunAsAdmin "New-Item `"$Path`" -Force | New-ItemProperty -Name `"$Name`" -Value (ConvertFrom-Json `"$Json`") -PropertyType $Type -Force"
    } elseif (-not ((ConvertTo-Json (Get-ItemProperty $Path | Select-Object -ExpandProperty $Name -ErrorAction Ignore)) -eq $Json)) {
        RunAsAdmin "Set-ItemProperty `"$Path`" -Name `"$Name`" -Value (ConvertFrom-Json `"$Json`") -Type $Type -Force"
    }
}

function RunAsAdmin {
    param($Command)
    $bytes = [System.Text.Encoding]::Unicode.GetBytes($Command)
    $encodedCommand = [Convert]::ToBase64String($bytes)
    Start-Process powershell -Verb runAs -ArgumentList "-EncodedCommand $encodedCommand"
}

function UnpackUrl {
    param($Url,$File,$UnpackDir,$TestDir)
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
                $shell = New-Object -com shell.application
                $shell.Namespace($UnpackDir).CopyHere($shell.Namespace($Output).Items())
            }
            '.7z' {
                & "C:\Program Files\7-Zip\7z.exe" x "-o$UnpackDir" "$Output" | Out-Null
            }
        }
    }
}

# disable bits branchcache https://powershell.org/forums/topic/bits-transfer-with-github/
Registry -Path HKLM:\SOFTWARE\Policies\Microsoft\Windows\BITS -Name DisableBranchCache -Value 1 -Type DWord

# swap capslock ctrl
Registry -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Keyboard Layout" -Name "Scancode Map" -Value 0,0,0,0,0,0,0,0,3,0,0,0,29,0,58,0,58,0,29,0,0,0,0,0 -Type Binary

# disable suggested apps
Registry -Path HKLM:\SOFTWARE\Policies\Microsoft\Windows\CloudContent -Name DisableWindowsConsumerFeatures -Value 1 -Type DWord
Registry -Path HKCU:\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager -Name SystemPaneSuggestionsEnabled -Value 0 -Type DWord

# install 7z
InstallUrl -DisplayName 7-Zip -Url "https://www.7-zip.org/a/7z1805-x64.exe" -Arg /S

# install firefox
InstallUrl -DisplayName Firefox -Url "https://download.mozilla.org/?product=firefox-latest&os=win64&lang=en-US" -Arg /S

# install mpv
UnpackUrl -Url "https://cfhcable.dl.sourceforge.net/project/mpv-player-windows/64bit/mpv-x86_64-20180721-git-08a6827.7z" -UnpackDir "$Programs\mpv"
if (-not (Test-Path "$Env:ProgramData\Microsoft\Windows\Start Menu\Programs\mpv.lnk")) {
    RunAsAdmin "$Programs\mpv\installer\mpv-install.bat"
}

# install weasel
InstallUrl -DisplayName С�Ǻ�ݔ�뷨 -Url "https://dl.bintray.com/rime/weasel/weasel-0.11.1.0-installer.exe" -Arg /S

# enable wsl
if (-not (Test-Path "C:\Windows\System32\wsl.exe")) {
    RunAsAdmin "Enable-WindowsOptionalFeature -Online -FeatureName Microsoft-Windows-Subsystem-Linux"
}

# install debian
if (-not (Get-AppxPackage -Name TheDebianProject.DebianGNULinux)) {
    Start-Process "ms-windows-store://pdp/?ProductId=9MSVKQC78PK6"
}

# install qterminal
UnpackUrl -Url "https://github.com/kghost/qterminal/releases/download/0.9.0-wsl.1/QTerminal.X64.zip" -UnpackDir "$Programs" -TestDir "$Programs\QTerminal"
CreateShortcut -Shortcut "$Desktop\QTerminal.lnk" -TargetPath "$Programs\QTerminal\QTerminal.exe"

# auto start sshd
CreateShortcut -Shortcut "$Startup\sshd.lnk" -TargetPath C:\Windows\System32\wsl.exe -Arguments "sudo service ssh start" -WindowStyle 7
if (-not (Get-Process sshd -ErrorAction Ignore)) {
    Invoke-Item -Path $sshd
}

# allow sshd firewall inbound
FirewallRule -DisplayName "WSL OpenSSH Server" -LocalPort 22

# whitelist wsl in windows defender
DefenderExcludeAppx -Name TheDebianProject.DebianGNULinux
