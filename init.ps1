$Desktop = [Environment]::GetFolderPath("Desktop")
$Startup = [Environment]::GetFolderPath("Startup")
$Programs = "$Env:HOMEDRIVE:\Programs"

function MakeDir {
    param($Dir)
    if (-not (Test-Path $Dir)) {
        mkdir -Path $Dir | Out-Null
    }
}

# create programs dir
MakeDir -Dir $Programs

function CreateShortcut {
    param($Shortcut,$TargetPath,$Arguments,$WindowStyle)
    if (-not (Test-Path $Shortcut)) {
        Write-Host "CreateShortcut: $Shortcut -> $TargetPath"
        $ws = New-Object -ComObject ("WScript.Shell")
        $sc = $ws.CreateShortcut($Shortcut)
        $sc.TargetPath = $TargetPath
        $sc.WorkingDirectory = (Get-Item $TargetPath).Directory.FullName
        if (-not ($Arguments -eq $null)) { $sc.Arguments = $Arguments }
        if (-not ($WindowStyle -eq $null)) { $sc.WindowStyle = $WindowStyle }
        $sc.Save()
    }
}

function EnableWslService {
    param($Service,$Process)
    if ($Process -eq $null) {
        $Process = $Service
    }
    CreateShortcut -Shortcut "$Startup\wsl-$Service.lnk" -TargetPath C:\Windows\System32\wsl.exe -Arguments "sudo service $Service start" -WindowStyle 7
    if (-not (Get-Process $Process -ErrorAction Ignore)) {
        Invoke-Item -Path "$Startup\wsl-$Service.lnk"
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
        Write-Host "DefenderExcludePath: $Path"
        RunAsAdmin "Set-MpPreference -ExclusionPath `"$Path`""
    }
}

function FirewallRule {
    param($DisplayName,$Action = "Allow",$Protocol = "TCP",$LocalPort)
    if (-not (Get-NetFirewallRule -DisplayName $DisplayName -ErrorAction Ignore)) {
        Write-Host "FirewallRule: $Action $Protocol $LocalPort"
        RunAsAdmin "New-NetFirewallRule -DisplayName `"$DisplayName`" -Action `"$Action`" -Protocol `"$Protocol`" -LocalPort `"$LocalPort`""
    }
}

# Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*,HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Select-Object DisplayName
function InstallUrl {
    param($DisplayName,$Url,$File,$Arg)
    if ($File -eq $null) {
        $File = $Url.Substring($Url.LastIndexOf("/") + 1)
    }
    $output = "$Home\Downloads\$File"
    if (-not (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*,HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object { $_.DisplayName -match $DisplayName })) {
        if (-not (Test-Path $output)) {
            Import-Module BitsTransfer
            Start-BitsTransfer -Description "Downloading $DisplayName installer from $Url" -Source $Url -Destination $output
        }
        Write-Host "InstallUrl: $DisplayName"
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
    $command = ""
    if (-not (Test-Path $Path)) {
        $command = "New-Item `"$Path`" -Force | New-ItemProperty -Name `"$Name`" -PropertyType $Type -Force -Value "
    } elseif (-not ((ConvertTo-Json (Get-ItemProperty $Path | Select-Object -ExpandProperty $Name -ErrorAction Ignore)) -eq $Json)) {
        $command = "Set-ItemProperty `"$Path`" -Name `"$Name`" -Type $Type -Force -Value "
    }
    if (-not ($command -eq "")) {
        Write-Host "Registry: $Path!$Name -> $Value"
        if ($Type -eq "String") {
            $command += "`"$Value`""
        } else {
            $command += "(ConvertFrom-Json `"$Json`")"
        }
        RunAsAdmin $command
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
        Write-Host "UnpackUrl: $Url -> $UnpackDir"
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

# remove shortcut arrow
Registry -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Icons" -Name 29 -Value "%SystemRoot%\System32\imageres.dll,-1015" -Type String
Start-Process ie4uinit.exe -ArgumentList -show

# Change UI font of some apps, default is Simsun for CN lang
Registry -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\GRE_Initialize" -Name "GUIFont.Facename" -Value "Segoe UI, Microsoft Yahei UI" -Type String

# enable remote app
Registry -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services" -Name fAllowUnlistedRemotePrograms -Value 1 -Type DWord

# disable suggested apps
Registry -Path HKLM:\SOFTWARE\Policies\Microsoft\Windows\CloudContent -Name DisableWindowsConsumerFeatures -Value 1 -Type DWord
Registry -Path HKCU:\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager -Name SystemPaneSuggestionsEnabled -Value 0 -Type DWord

# install vcredist 2010 x86 for youtube-dl
InstallUrl -DisplayName "Microsoft Visual C\+\+ 2010  x86 Redistributable" -Url "https://download.microsoft.com/download/5/B/C/5BC5DBB3-652D-4DCE-B14A-475AB85EEF6E/vcredist_x86.exe" -File vcredist-2010-x86.exe -Arg "/q /norestart"

# install 7z
InstallUrl -DisplayName 7-Zip -Url "https://www.7-zip.org/a/7z1805-x64.exe" -Arg /S

# install firefox
InstallUrl -DisplayName Firefox -Url "https://download.mozilla.org/?product=firefox-latest&os=win64&lang=en-US" -File firefox-installer.exe -Arg /S

# install mpv
UnpackUrl -Url "https://cfhcable.dl.sourceforge.net/project/mpv-player-windows/64bit/mpv-x86_64-20190407-git-e9fae41.7z" -UnpackDir "$Programs\mpv"
if (-not (Test-Path "$Env:ProgramData\Microsoft\Windows\Start Menu\Programs\mpv.lnk")) {
    RunAsAdmin "$Programs\mpv\installer\mpv-install.bat"
}
MakeDir -Dir "$Programs\mpv\mpv\script-opts"
Set-Content -Path "$Programs\mpv\mpv\mpv.conf" -Force -Value @'
[default]
hwdec=auto
keep-open-pause=no
keep-open=yes
sub-auto=fuzzy
sub-codepage=gbk
user-agent="Mozilla/5.0 (iPad; CPU OS 8_1_3 like Mac OS X) AppleWebKit/600.1.4 (KHTML, like Gecko) Version/8.0 Mobile/12B466 Safari/600.1.4"
'@
Set-Content -Path "$Programs\mpv\mpv\input.conf" -Force -Value @'
Y add sub-scale +0.1                # increase subtitle font size
G add sub-scale -0.1                # decrease subtitle font size
y sub_step -1                       # immediately display next subtitle
g sub_step +1                       # previous
R cycle_values window-scale 2 0.5 1 # switch between 2x, 1/2, unresized window size
'@
Set-Content -Path "$Programs\mpv\mpv\script-opts\osc.conf" -Force -Value @'
seekbarstyle=knob
deadzonesize=1
minmousemove=1
'@

# install weasel
InstallUrl -DisplayName С�Ǻ�ݔ�뷨 -Url "https://dl.bintray.com/rime/weasel/weasel-0.13.0.0-installer.exe" -Arg /S

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
EnableWslService -Service ssh -Process sshd

# allow sshd firewall inbound
FirewallRule -DisplayName "WSL OpenSSH Server" -LocalPort 22

# auto start cron
EnableWslService -Service cron

# whitelist wsl in windows defender
DefenderExcludeAppx -Name TheDebianProject.DebianGNULinux

# install VcXsrv
InstallUrl -DisplayName VcXsrv -Url https://cfhcable.dl.sourceforge.net/project/vcxsrv/vcxsrv/1.20.1.4/vcxsrv-64.1.20.1.4.installer.exe -Arg /S

# misc notes
# disable vpn default route
# Set-VpnConnection -Name "<vpnname>" -SplitTunneling $True
