#Requires -Version 5.1
<#
.SYNOPSIS
    Windows 11 環境を一から再構築するための設定情報を網羅的に収集する

.DESCRIPTION
    システム設定・インストール済みソフト・ネットワーク・セキュリティ等を調査し、
    カテゴリごとにファイルへ保存する。
    管理者権限で実行すると収集できる情報が増える。

.PARAMETER OutputPath
    出力先ディレクトリ（省略時はスクリプトと同じ場所）

.EXAMPLE
    .\Collect-WindowsEnv.ps1
    .\Collect-WindowsEnv.ps1 -OutputPath "D:\EnvBackup"
#>
param(
    [string]$OutputPath = $PSScriptRoot
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'SilentlyContinue'

# 日本語Windowsの外部コマンド出力（CP932/Shift-JIS）を正しく読み取る
# PS 5.1 ではデフォルトが CP932 なので変更不要だが、PS 7+ は UTF-8 がデフォルトのため明示設定する
[Console]::OutputEncoding = [System.Text.Encoding]::GetEncoding(932)

# ─────────────────────────────────────────────
# 初期化
# ─────────────────────────────────────────────
$timestamp  = Get-Date -Format 'yyyyMMdd_HHmmss'
$outDir     = Join-Path $OutputPath "WindowsEnvAudit_$timestamp"
New-Item -ItemType Directory -Path $outDir -Force | Out-Null

$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
    [Security.Principal.WindowsBuiltInRole]::Administrator)

$indexLines = @("# Windows 11 環境調査レポート", "生成日時: $(Get-Date)", "実行ユーザー: $env:USERNAME", "管理者権限: $isAdmin", "出力先: $outDir", "")
$progress   = 0
$total      = 20

function Write-Step {
    param([int]$n, [string]$name)
    $script:progress = $n
    Write-Progress -Activity "環境調査中" -Status "[$n/$total] $name" -PercentComplete ([int]($n / $total * 100))
    Write-Host "[$n/$total] $name ..." -ForegroundColor Cyan
}

function Save-Json {
    param([string]$filename, $data)
    $path = Join-Path $outDir $filename
    $data | ConvertTo-Json -Depth 10 | Set-Content -Path $path -Encoding UTF8
    return $path
}

function Save-Text {
    param([string]$filename, [string[]]$lines)
    $path = Join-Path $outDir $filename
    $lines | Set-Content -Path $path -Encoding UTF8
    return $path
}

function Append-Index {
    param([string]$category, [string]$file)
    $script:indexLines += "## $category"
    $script:indexLines += "  ファイル: $(Split-Path $file -Leaf)"
    $script:indexLines += ""
}

function Try-Command {
    param([scriptblock]$sb, $fallback = $null)
    try { & $sb } catch { $fallback }
}

# ─────────────────────────────────────────────
# 01. システム基本情報
# ─────────────────────────────────────────────
Write-Step 1 "システム基本情報"

$osInfo  = Get-CimInstance Win32_OperatingSystem
$csInfo  = Get-CimInstance Win32_ComputerSystem
$cpuInfo = Get-CimInstance Win32_Processor
$gpuInfo = Get-CimInstance Win32_VideoController
$biosInfo= Get-CimInstance Win32_BIOS
$mbInfo  = Get-CimInstance Win32_BaseBoard

# セキュアブート状態（Confirm-SecureBootUEFI は管理者不要だが環境依存）
$secureBoot = Try-Command { Confirm-SecureBootUEFI } "不明"

$systemInfo = [ordered]@{
    OS = [ordered]@{
        Caption       = $osInfo.Caption
        Version       = $osInfo.Version
        BuildNumber   = $osInfo.BuildNumber
        OSArchitecture= $osInfo.OSArchitecture
        InstallDate   = $osInfo.InstallDate
        LastBootUpTime= $osInfo.LastBootUpTime
        WindowsDirectory = $osInfo.WindowsDirectory
        SystemDrive   = $osInfo.SystemDrive
    }
    Computer = [ordered]@{
        Name           = $csInfo.Name
        Domain         = $csInfo.Domain
        Workgroup      = $csInfo.Workgroup
        Manufacturer   = $csInfo.Manufacturer
        Model          = $csInfo.Model
        TotalPhysicalMemoryGB = [math]::Round($csInfo.TotalPhysicalMemory / 1GB, 2)
    }
    CPU = @($cpuInfo | ForEach-Object {
        [ordered]@{
            Name          = $_.Name
            NumberOfCores = $_.NumberOfCores
            NumberOfLogicalProcessors = $_.NumberOfLogicalProcessors
            MaxClockSpeedMHz = $_.MaxClockSpeed
        }
    })
    GPU = @($gpuInfo | ForEach-Object {
        [ordered]@{
            Name             = $_.Name
            AdapterRAMGB     = [math]::Round($_.AdapterRAM / 1GB, 2)
            DriverVersion    = $_.DriverVersion
            VideoModeDescription = $_.VideoModeDescription
        }
    })
    BIOS = [ordered]@{
        Manufacturer  = $biosInfo.Manufacturer
        Version       = $biosInfo.Version
        SMBIOSVersion = "$($biosInfo.SMBIOSMajorVersion).$($biosInfo.SMBIOSMinorVersion)"
        ReleaseDate   = $biosInfo.ReleaseDate
        SecureBoot    = $secureBoot
    }
    Motherboard = [ordered]@{
        Manufacturer = $mbInfo.Manufacturer
        Product      = $mbInfo.Product
        SerialNumber = $mbInfo.SerialNumber
    }
}
$f = Save-Json "01_system_info.json" $systemInfo
Append-Index "01. システム基本情報" $f

# ─────────────────────────────────────────────
# 02. ディスク・パーティション
# ─────────────────────────────────────────────
Write-Step 2 "ディスク・パーティション"

$disks = Get-Disk | ForEach-Object {
    $disk = $_
    $parts = Get-Partition -DiskNumber $disk.Number -ErrorAction SilentlyContinue
    [ordered]@{
        DiskNumber       = $disk.Number
        FriendlyName     = $disk.FriendlyName
        SizeGB           = [math]::Round($disk.Size / 1GB, 2)
        PartitionStyle   = $disk.PartitionStyle
        BusType          = $disk.BusType
        MediaType        = $disk.MediaType
        OperationalStatus= $disk.OperationalStatus
        Partitions       = @($parts | ForEach-Object {
            $vol = Get-Volume -Partition $_ -ErrorAction SilentlyContinue
            [ordered]@{
                PartitionNumber = $_.PartitionNumber
                Type            = $_.Type
                SizeGB          = [math]::Round($_.Size / 1GB, 2)
                DriveLetter     = $_.DriveLetter
                IsSystem        = $_.IsSystem
                IsBoot          = $_.IsBoot
                FileSystem      = $vol.FileSystem
                Label           = $vol.FileSystemLabel
            }
        })
    }
}
$f = Save-Json "02_disk_partitions.json" $disks
Append-Index "02. ディスク・パーティション" $f

# ─────────────────────────────────────────────
# 03. インストール済みアプリ（Win32）
# ─────────────────────────────────────────────
Write-Step 3 "インストール済みアプリ（Win32）"

$regPaths = @(
    'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*',
    'HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*',
    'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*'
)
$apps = $regPaths | ForEach-Object {
    Get-ItemProperty $_ -ErrorAction SilentlyContinue
} | Where-Object { $_.DisplayName } |
  Select-Object DisplayName, DisplayVersion, Publisher, InstallDate, InstallLocation, UninstallString |
  Sort-Object DisplayName |
  ForEach-Object {
    [ordered]@{
        Name            = $_.DisplayName
        Version         = $_.DisplayVersion
        Publisher       = $_.Publisher
        InstallDate     = $_.InstallDate
        InstallLocation = $_.InstallLocation
        UninstallString = $_.UninstallString
    }
}
$f = Save-Json "03_installed_apps_win32.json" $apps
Append-Index "03. インストール済みアプリ（Win32）" $f

# MSI インストーラー直接インストール（ProductCode 等の詳細情報）
$msiApps = Get-Package -ProviderName msi -ErrorAction SilentlyContinue |
    Sort-Object Name | ForEach-Object {
        [ordered]@{
            Name        = $_.Name
            Version     = $_.Version
            ProductCode = $_.Metadata['ProductCode']
            UpgradeCode = $_.Metadata['UpgradeCode']
            Source      = $_.Source
        }
    }
$f = Save-Json "03b_installed_apps_msi.json" $msiApps
Append-Index "03b. インストール済みアプリ（MSI直接インストール）" $f

# ポータブルアプリスキャン（インストーラーなしで配置されたEXEの検出）
Write-Host "  -> ポータブルアプリをスキャン中..." -ForegroundColor Gray

# Win32 登録済みインストール先パスを照合用に収集（小文字・末尾スラッシュ除去）
$registeredLocations = @(
    $regPaths | ForEach-Object {
        Get-ItemProperty $_ -ErrorAction SilentlyContinue
    } | Where-Object { $_.InstallLocation } |
      ForEach-Object { $_.InstallLocation.TrimEnd('\').ToLower() } |
      Where-Object { $_ }
)

# スキャン対象ディレクトリと深さ定義
$scanTargets = @(
    @{ Path = $env:ProgramFiles;                Depth = 3 }
    @{ Path = ${env:ProgramFiles(x86)};         Depth = 3 }
    @{ Path = "$env:LOCALAPPDATA\Programs";     Depth = 3 }
    @{ Path = "$env:USERPROFILE\Desktop";       Depth = 2 }
    @{ Path = "$env:USERPROFILE\Downloads";     Depth = 2 }
    @{ Path = "$env:USERPROFILE\Documents";     Depth = 2 }
    @{ Path = "$env:USERPROFILE\PortableApps";  Depth = 3 }
)

$portableScan = @()
foreach ($target in $scanTargets) {
    if (-not (Test-Path $target.Path)) { continue }

    $exeFiles = Get-ChildItem -Path $target.Path -Filter '*.exe' `
        -Recurse -Depth $target.Depth -ErrorAction SilentlyContinue |
        Where-Object { $_.FullName -notmatch '^C:\\Windows\\' }

    foreach ($exe in $exeFiles) {
        $vi     = $exe.VersionInfo
        $exeDir = $exe.DirectoryName.ToLower()
        $isReg  = ($registeredLocations | Where-Object { $_ -and $exeDir.StartsWith($_) }).Count -gt 0

        $portableScan += [ordered]@{
            FullPath     = $exe.FullName
            ProductName  = $vi.ProductName
            FileVersion  = $vi.FileVersion
            CompanyName  = $vi.CompanyName
            Description  = $vi.FileDescription
            LastModified = $exe.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss')
            IsRegistered = $isReg
        }
    }
}

$f = Save-Json "03c_portable_apps_scan.json" $portableScan
Append-Index "03c. ポータブルアプリスキャン（EXE検出）" $f

# ─────────────────────────────────────────────
# 04. Store アプリ (AppX/MSIX)
# ─────────────────────────────────────────────
Write-Step 4 "Store アプリ (AppX/MSIX)"

$storeApps = Get-AppxPackage -AllUsers | Select-Object Name, Version, Publisher, Architecture, PackageFullName |
    Sort-Object Name | ForEach-Object {
        [ordered]@{
            Name            = $_.Name
            Version         = $_.Version
            Publisher       = $_.Publisher
            Architecture    = $_.Architecture
            PackageFullName = $_.PackageFullName
        }
    }
$f = Save-Json "04_installed_apps_store.json" $storeApps
Append-Index "04. Store アプリ (AppX/MSIX)" $f

# ─────────────────────────────────────────────
# 05. パッケージマネージャー (winget / choco / scoop)
# ─────────────────────────────────────────────
Write-Step 5 "パッケージマネージャー"

$pkgLines = @()

# winget
$wingetCmd  = Get-Command winget -ErrorAction SilentlyContinue
$wingetPath = if ($wingetCmd) { $wingetCmd.Source } else { $null }
if ($wingetPath) {
    $pkgLines += "=== winget list ==="
    # winget は UTF-8 で出力するため一時的に OutputEncoding を切り替える
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
    $pkgLines += (winget list --accept-source-agreements 2>&1)
    [Console]::OutputEncoding = [System.Text.Encoding]::GetEncoding(932)
    $pkgLines += ""
} else {
    $pkgLines += "=== winget: 未インストール ==="
    $pkgLines += ""
}

# Chocolatey
if (Get-Command choco -ErrorAction SilentlyContinue) {
    $pkgLines += "=== Chocolatey list ==="
    $pkgLines += (choco list 2>&1)
    $pkgLines += ""
} else {
    $pkgLines += "=== Chocolatey: 未インストール ==="
    $pkgLines += ""
}

# Scoop
if (Get-Command scoop -ErrorAction SilentlyContinue) {
    $pkgLines += "=== Scoop list ==="
    $pkgLines += (scoop list 2>&1)
    $pkgLines += ""
} else {
    $pkgLines += "=== Scoop: 未インストール ==="
    $pkgLines += ""
}

$f = Save-Text "05_package_managers.txt" $pkgLines
Append-Index "05. パッケージマネージャー" $f

# ─────────────────────────────────────────────
# 06. 開発ツール・ランタイム
# ─────────────────────────────────────────────
Write-Step 6 "開発ツール・ランタイム"

function Get-ToolVersion {
    param([string]$cmd, [string]$args = "--version")
    $cmdInfo = Get-Command $cmd -ErrorAction SilentlyContinue
    if ($cmdInfo -and $cmdInfo.Source -notlike "*WindowsApps*") {
        $v = (& $cmd $args 2>&1) | Select-Object -First 1
        return $v -replace "`r`n|`n", ""
    }
    return $null
}

$devTools = [ordered]@{
    DotNet = [ordered]@{
        dotnet        = Get-ToolVersion "dotnet" "--version"
        SDKs          = Try-Command { (dotnet --list-sdks 2>&1) } $null
        Runtimes      = Try-Command { (dotnet --list-runtimes 2>&1) } $null
    }
    Node = [ordered]@{
        node          = Get-ToolVersion "node"
        npm           = Get-ToolVersion "npm"
        GlobalPackages= Try-Command { (npm list -g --depth=0 2>&1) } $null
    }
    Python = [ordered]@{
        python        = Get-ToolVersion "python"
        python3       = Get-ToolVersion "python3"
        pip           = Get-ToolVersion "pip"
        PipPackages   = Try-Command { (pip list 2>&1) } $null
    }
    Java = [ordered]@{
        java          = Get-ToolVersion "java"
        javac         = Get-ToolVersion "javac"
        JAVA_HOME     = $env:JAVA_HOME
    }
    Go = [ordered]@{
        go            = Get-ToolVersion "go" "version"
        GOPATH        = $env:GOPATH
        GOROOT        = $env:GOROOT
    }
    Rust = [ordered]@{
        rustc         = Get-ToolVersion "rustc"
        cargo         = Get-ToolVersion "cargo"
    }
    Ruby = [ordered]@{
        ruby          = Get-ToolVersion "ruby"
        gem           = Get-ToolVersion "gem"
    }
    Git = [ordered]@{
        version       = Get-ToolVersion "git"
        globalConfig  = Try-Command { (git config --global --list 2>&1) } $null
    }
    Docker = [ordered]@{
        docker        = Get-ToolVersion "docker"
        compose       = Get-ToolVersion "docker" "compose version"
    }
    PowerShell = [ordered]@{
        PSVersion     = $PSVersionTable.PSVersion.ToString()
        Edition       = $PSVersionTable.PSEdition
        InstalledVersions = @(
            Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\PowerShellCore\InstalledVersions\*' -ErrorAction SilentlyContinue |
                Select-Object SemanticVersion, InstallDate
        )
    }
    OtherTools = [ordered]@{
        curl          = Get-ToolVersion "curl"
        wget          = Get-ToolVersion "wget"
        make          = Get-ToolVersion "make"
        cmake         = Get-ToolVersion "cmake"
        kubectl       = Get-ToolVersion "kubectl" "version --client --short"
        helm          = Get-ToolVersion "helm" "version --short"
        terraform     = Get-ToolVersion "terraform"
        aws           = Get-ToolVersion "aws" "--version"
        az            = Get-ToolVersion "az" "version"
    }
}
$f = Save-Json "06_dev_tools.json" $devTools
Append-Index "06. 開発ツール・ランタイム" $f

# VSCode 拡張機能
if (Get-Command code -ErrorAction SilentlyContinue) {
    $vscodeExt = (code --list-extensions --show-versions 2>&1)
    $f = Save-Text "06b_vscode_extensions.txt" $vscodeExt
    Append-Index "06b. VSCode 拡張機能" $f
}

# WSL
if (Get-Command wsl -ErrorAction SilentlyContinue) {
    $wslInfo = @()
    $wslInfo += "=== WSL --list --verbose ==="
    $wslInfo += (wsl --list --verbose 2>&1)
    $wslInfo += ""
    $wslInfo += "=== WSL --status ==="
    $wslInfo += (wsl --status 2>&1)
    $f = Save-Text "06c_wsl.txt" $wslInfo
    Append-Index "06c. WSL ディストリビューション" $f
}

# PowerShell モジュール
$psModules = Get-Module -ListAvailable | Select-Object Name, Version, ModuleType, Path |
    Sort-Object Name | ForEach-Object {
        [ordered]@{ Name=$_.Name; Version=$_.Version.ToString(); ModuleType=$_.ModuleType; Path=$_.Path }
    }
$f = Save-Json "06d_powershell_modules.json" $psModules
Append-Index "06d. PowerShell モジュール" $f

# ─────────────────────────────────────────────
# 07. Windows オプション機能
# ─────────────────────────────────────────────
Write-Step 7 "Windows オプション機能"

$optFeatures = Get-WindowsOptionalFeature -Online -ErrorAction SilentlyContinue |
    Select-Object FeatureName, State | Sort-Object FeatureName |
    ForEach-Object { [ordered]@{ Feature=$_.FeatureName; State=$_.State.ToString() } }
$f = Save-Json "07_windows_optional_features.json" $optFeatures
Append-Index "07. Windows オプション機能" $f

# Windows Capability
$capabilities = Get-WindowsCapability -Online -ErrorAction SilentlyContinue |
    Select-Object Name, State | Sort-Object Name |
    ForEach-Object { [ordered]@{ Name=$_.Name; State=$_.State.ToString() } }
$f = Save-Json "07b_windows_capabilities.json" $capabilities
Append-Index "07b. Windows Capabilities" $f

# ─────────────────────────────────────────────
# 08. サービス
# ─────────────────────────────────────────────
Write-Step 8 "サービス"

$services = Get-Service | ForEach-Object {
    $wmiSvc = Get-CimInstance Win32_Service -Filter "Name='$($_.Name)'" -ErrorAction SilentlyContinue
    [ordered]@{
        Name        = $_.Name
        DisplayName = $_.DisplayName
        Status      = $_.Status.ToString()
        StartType   = $_.StartType.ToString()
        PathName    = $wmiSvc.PathName
        Description = $wmiSvc.Description
        StartName   = $wmiSvc.StartName
    }
} | Sort-Object Name
$f = Save-Json "08_services.json" $services
Append-Index "08. サービス" $f

# ─────────────────────────────────────────────
# 09. ネットワーク設定
# ─────────────────────────────────────────────
Write-Step 9 "ネットワーク設定"

# アダプターと IP 設定
$netAdapters = Get-NetAdapter | Where-Object { $_.Status -ne 'Not Present' } | ForEach-Object {
    $adapter = $_
    $ipConfig = Get-NetIPConfiguration -InterfaceIndex $adapter.InterfaceIndex -ErrorAction SilentlyContinue
    $dnsServers = (Get-DnsClientServerAddress -InterfaceIndex $adapter.InterfaceIndex -AddressFamily IPv4 -ErrorAction SilentlyContinue).ServerAddresses
    [ordered]@{
        Name            = $adapter.Name
        InterfaceAlias  = $adapter.InterfaceAlias
        InterfaceDescription = $adapter.InterfaceDescription
        Status          = $adapter.Status.ToString()
        MacAddress      = $adapter.MacAddress
        LinkSpeedMbps   = [math]::Round($adapter.LinkSpeed / 1MB, 0)
        IPv4Address     = @($ipConfig.IPv4Address.IPAddress)
        IPv4PrefixLength= @($ipConfig.IPv4Address.PrefixLength)
        DefaultGateway  = @($ipConfig.IPv4DefaultGateway.NextHop)
        DNSServers      = @($dnsServers)
        IPv6Address     = @($ipConfig.IPv6Address.IPAddress)
    }
}
$f = Save-Json "09_network_adapters.json" $netAdapters
Append-Index "09. ネットワークアダプター" $f

# hosts ファイル
$hostsPath = "$env:SystemRoot\System32\drivers\etc\hosts"
$f = Save-Text "09b_hosts_file.txt" (Get-Content $hostsPath -ErrorAction SilentlyContinue)
Append-Index "09b. hosts ファイル" $f

# プロキシ設定
$proxyReg = Get-ItemProperty 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings' -ErrorAction SilentlyContinue
$proxy = [ordered]@{
    ProxyEnabled  = $proxyReg.ProxyEnable
    ProxyServer   = $proxyReg.ProxyServer
    ProxyOverride = $proxyReg.ProxyOverride
    AutoConfigURL = $proxyReg.AutoConfigURL
}
$f = Save-Json "09c_proxy_settings.json" $proxy
Append-Index "09c. プロキシ設定" $f

# Wi-Fi プロファイル
$wifiProfiles = @()
$profileNames = (netsh wlan show profiles 2>&1) | Select-String "All User Profile" |
    ForEach-Object { ($_ -split ':')[-1].Trim() }
foreach ($name in $profileNames) {
    $wifiProfiles += "=== $name ==="
    $wifiProfiles += (netsh wlan show profile name="$name" 2>&1 | Select-String -NotMatch "Key Content")
    $wifiProfiles += ""
}
$f = Save-Text "09d_wifi_profiles.txt" $wifiProfiles
Append-Index "09d. Wi-Fi プロファイル" $f

# ファイアウォールルール（カスタムのみ）
if ($isAdmin) {
    $fwRules = Get-NetFirewallRule | Where-Object { $_.PolicyStoreSourceType -eq 'Local' } |
        Select-Object Name, DisplayName, Direction, Action, Enabled, Profile, Protocol |
        ForEach-Object {
            [ordered]@{
                Name        = $_.Name
                DisplayName = $_.DisplayName
                Direction   = $_.Direction.ToString()
                Action      = $_.Action.ToString()
                Enabled     = $_.Enabled.ToString()
                Profile     = $_.Profile.ToString()
                Protocol    = $_.Protocol
            }
        }
    $f = Save-Json "09e_firewall_rules_custom.json" $fwRules
    Append-Index "09e. ファイアウォールカスタムルール" $f
}

# ─────────────────────────────────────────────
# 10. ユーザー・セキュリティ設定
# ─────────────────────────────────────────────
Write-Step 10 "ユーザー・セキュリティ設定"

# ローカルユーザー
$localUsers = Get-LocalUser | ForEach-Object {
    [ordered]@{
        Name             = $_.Name
        Enabled          = $_.Enabled
        Description      = $_.Description
        FullName         = $_.FullName
        PasswordRequired = $_.PasswordRequired
        PasswordExpires  = $_.PasswordExpires
        LastLogon        = $_.LastLogon
        SID              = $_.SID.Value
    }
}
$f = Save-Json "10_local_users.json" $localUsers
Append-Index "10. ローカルユーザー" $f

# ローカルグループ
$localGroups = Get-LocalGroup | ForEach-Object {
    $grp = $_
    $members = Try-Command { (Get-LocalGroupMember -Group $grp.Name -ErrorAction SilentlyContinue | Select-Object Name, ObjectClass, PrincipalSource) } @()
    [ordered]@{
        Name        = $grp.Name
        Description = $grp.Description
        SID         = $grp.SID.Value
        Members     = @($members | ForEach-Object { [ordered]@{ Name=$_.Name; Type=$_.ObjectClass } })
    }
}
$f = Save-Json "10b_local_groups.json" $localGroups
Append-Index "10b. ローカルグループ" $f

# 環境変数
$envVars = [ordered]@{
    System = [System.Environment]::GetEnvironmentVariables([System.EnvironmentVariableTarget]::Machine).GetEnumerator() |
        Sort-Object Key | ForEach-Object { [ordered]@{ Name=$_.Key; Value=$_.Value } }
    User   = [System.Environment]::GetEnvironmentVariables([System.EnvironmentVariableTarget]::User).GetEnumerator() |
        Sort-Object Key | ForEach-Object { [ordered]@{ Name=$_.Key; Value=$_.Value } }
}
$f = Save-Json "10c_environment_variables.json" $envVars
Append-Index "10c. 環境変数" $f

# UAC 設定
$uacReg = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System' -ErrorAction SilentlyContinue
$uac = [ordered]@{
    EnableLUA             = $uacReg.EnableLUA
    ConsentPromptBehaviorAdmin  = $uacReg.ConsentPromptBehaviorAdmin
    ConsentPromptBehaviorUser   = $uacReg.ConsentPromptBehaviorUser
    PromptOnSecureDesktop= $uacReg.PromptOnSecureDesktop
}
$f = Save-Json "10d_uac_settings.json" $uac
Append-Index "10d. UAC 設定" $f

# BitLocker
if ($isAdmin) {
    $bitlocker = Get-BitLockerVolume -ErrorAction SilentlyContinue | ForEach-Object {
        [ordered]@{
            MountPoint           = $_.MountPoint
            VolumeStatus         = $_.VolumeStatus.ToString()
            ProtectionStatus     = $_.ProtectionStatus.ToString()
            EncryptionMethod     = $_.EncryptionMethod.ToString()
            EncryptionPercentage = $_.EncryptionPercentage
        }
    }
    $f = Save-Json "10e_bitlocker.json" $bitlocker
    Append-Index "10e. BitLocker" $f
}

# ─────────────────────────────────────────────
# 11. スタートアップ・スケジュールタスク
# ─────────────────────────────────────────────
Write-Step 11 "スタートアップ・スケジュールタスク"

# スタートアップレジストリ
$startupKeys = @(
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run',
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce',
    'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run',
    'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce'
)
$startupItems = $startupKeys | ForEach-Object {
    $key = $_
    $props = Get-ItemProperty $key -ErrorAction SilentlyContinue
    if ($props) {
        $props.PSObject.Properties | Where-Object { $_.Name -notmatch '^PS' } | ForEach-Object {
            [ordered]@{ Registry=$key; Name=$_.Name; Command=$_.Value }
        }
    }
}
$f = Save-Json "11_startup_registry.json" $startupItems
Append-Index "11. スタートアップ（レジストリ）" $f

# スタートアップフォルダ
$startupFolders = @(
    [System.Environment]::GetFolderPath('Startup'),
    [System.Environment]::GetFolderPath('CommonStartup')
)
$startupFiles = $startupFolders | ForEach-Object {
    $folder = $_
    if (Test-Path $folder) {
        Get-ChildItem $folder -Recurse | ForEach-Object {
            [ordered]@{ Folder=$folder; Name=$_.Name; FullPath=$_.FullName }
        }
    }
}
$f = Save-Json "11b_startup_folders.json" $startupFiles
Append-Index "11b. スタートアップフォルダ" $f

# スケジュールタスク（非 Microsoft）
$tasks = Get-ScheduledTask | Where-Object {
    $_.TaskPath -notlike '\Microsoft\*'
} | ForEach-Object {
    $info = Get-ScheduledTaskInfo -TaskName $_.TaskName -TaskPath $_.TaskPath -ErrorAction SilentlyContinue
    [ordered]@{
        TaskName    = $_.TaskName
        TaskPath    = $_.TaskPath
        State       = $_.State.ToString()
        Description = $_.Description
        LastRunTime = $info.LastRunTime
        NextRunTime = $info.NextRunTime
        Actions     = @($_.Actions | ForEach-Object { "$($_.Execute) $($_.Arguments)" })
        Triggers    = @($_.Triggers | ForEach-Object { $_.CimClass.CimClassName })
    }
}
$f = Save-Json "11c_scheduled_tasks.json" $tasks
Append-Index "11c. スケジュールタスク（非Microsoft）" $f

# ─────────────────────────────────────────────
# 12. 電源・パフォーマンス設定
# ─────────────────────────────────────────────
Write-Step 12 "電源・パフォーマンス設定"

$powerLines = @()
$powerLines += "=== アクティブな電源プラン ==="
$powerLines += (powercfg /getactivescheme 2>&1)
$powerLines += ""
$powerLines += "=== 全電源プラン ==="
$powerLines += (powercfg /list 2>&1)
$powerLines += ""
$powerLines += "=== 電源設定詳細 ==="
$powerLines += (powercfg /query 2>&1)
$f = Save-Text "12_power_settings.txt" $powerLines
Append-Index "12. 電源設定" $f

# 仮想メモリ
$compSys = Get-CimInstance Win32_ComputerSystem
$pageFile = Get-CimInstance Win32_PageFileUsage -ErrorAction SilentlyContinue
$virtMem = [ordered]@{
    AutomaticManagedPagefile = $compSys.AutomaticManagedPagefile
    PageFiles = @($pageFile | ForEach-Object {
        [ordered]@{
            Name               = $_.Name
            CurrentUsageMB     = $_.CurrentUsage
            AllocatedBaseSizeMB= $_.AllocatedBaseSize
            PeakUsageMB        = $_.PeakUsage
        }
    })
}
$f = Save-Json "12b_virtual_memory.json" $virtMem
Append-Index "12b. 仮想メモリ" $f

# ─────────────────────────────────────────────
# 13. 地域・言語・入力設定
# ─────────────────────────────────────────────
Write-Step 13 "地域・言語・入力設定"

$locale = [ordered]@{
    TimeZone       = (Get-TimeZone).Id
    TimeZoneDisplay= (Get-TimeZone).DisplayName
    Culture        = (Get-Culture).Name
    UICulture      = (Get-UICulture).Name
    SystemLocale   = (Get-WinSystemLocale).Name
    InputLanguages = @((Get-WinUserLanguageList) | ForEach-Object {
        [ordered]@{
            LanguageTag   = $_.LanguageTag
            Autonym       = $_.Autonym
            InputMethods  = @($_.InputMethodTips)
        }
    })
    HomeLocation   = (Get-WinHomeLocation).HomeLocation
    DateFormat     = (Get-Culture).DateTimeFormat.ShortDatePattern
    TimeFormat     = (Get-Culture).DateTimeFormat.ShortTimePattern
    FirstDayOfWeek = (Get-Culture).DateTimeFormat.FirstDayOfWeek.ToString()
}
$f = Save-Json "13_locale_language.json" $locale
Append-Index "13. 地域・言語・入力設定" $f

# ─────────────────────────────────────────────
# 14. ディスプレイ設定
# ─────────────────────────────────────────────
Write-Step 14 "ディスプレイ設定"

Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public class DisplayHelper {
    [DllImport("user32.dll")]
    public static extern bool EnumDisplaySettings(string deviceName, int modeNum, ref DEVMODE devMode);
    [StructLayout(LayoutKind.Sequential)]
    public struct DEVMODE {
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)]
        public string dmDeviceName;
        public short dmSpecVersion, dmDriverVersion, dmSize, dmDriverExtra;
        public int dmFields;
        public int dmPositionX, dmPositionY;
        public int dmDisplayOrientation, dmDisplayFixedOutput;
        public short dmColor, dmDuplex, dmYResolution, dmTTOption, dmCollate;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)]
        public string dmFormName;
        public short dmLogPixels;
        public int dmBitsPerPel, dmPelsWidth, dmPelsHeight, dmDisplayFlags, dmDisplayFrequency;
    }
}
"@ -ErrorAction SilentlyContinue

$monitors = Get-CimInstance Win32_VideoController | ForEach-Object {
    [ordered]@{
        Name                 = $_.Name
        CurrentHorizontalRes = $_.CurrentHorizontalResolution
        CurrentVerticalRes   = $_.CurrentVerticalResolution
        CurrentRefreshRate   = $_.CurrentRefreshRate
        BitsPerPixel         = $_.CurrentBitsPerPixel
        DriverVersion        = $_.DriverVersion
        DriverDate           = $_.DriverDate
    }
}

# DPI スケーリング（レジストリから）
$dpiReg = Get-ItemProperty 'HKCU:\Control Panel\Desktop' -ErrorAction SilentlyContinue
$displays = [ordered]@{
    Monitors   = @($monitors)
    LogPixels  = $dpiReg.LogPixels
    Win8DpiScaling = $dpiReg.Win8DpiScaling
}
$f = Save-Json "14_display_settings.json" $displays
Append-Index "14. ディスプレイ設定" $f

# ─────────────────────────────────────────────
# 15. 共有フォルダ・ネットワークドライブ
# ─────────────────────────────────────────────
Write-Step 15 "共有・ネットワークドライブ"

# 共有フォルダ
$shares = Get-SmbShare -ErrorAction SilentlyContinue | Where-Object { $_.Name -notmatch '\$$' } |
    ForEach-Object {
        [ordered]@{
            Name        = $_.Name
            Path        = $_.Path
            Description = $_.Description
            ShareState  = $_.ShareState.ToString()
        }
    }
$f = Save-Json "15_shared_folders.json" $shares
Append-Index "15. 共有フォルダ" $f

# ネットワークドライブ
$netDrives = Get-PSDrive -PSProvider FileSystem | Where-Object { $_.DisplayRoot -match '\\\\' } |
    ForEach-Object { [ordered]@{ Name=$_.Name; Root=$_.Root; DisplayRoot=$_.DisplayRoot } }
$f = Save-Json "15b_network_drives.json" $netDrives
Append-Index "15b. ネットワークドライブ" $f

# ─────────────────────────────────────────────
# 16. ドライバー
# ─────────────────────────────────────────────
Write-Step 16 "ドライバー"

$drivers = Get-CimInstance Win32_PnPSignedDriver |
    Select-Object DeviceName, DriverVersion, DriverDate, Manufacturer, IsSigned, InfName |
    Sort-Object DeviceName | ForEach-Object {
        [ordered]@{
            DeviceName    = $_.DeviceName
            DriverVersion = $_.DriverVersion
            DriverDate    = $_.DriverDate
            Manufacturer  = $_.Manufacturer
            IsSigned      = $_.IsSigned
            InfName       = $_.InfName
        }
    }
$f = Save-Json "16_drivers.json" $drivers
Append-Index "16. ドライバー" $f

# ─────────────────────────────────────────────
# 17. フォント・プリンター
# ─────────────────────────────────────────────
Write-Step 17 "フォント・プリンター"

# フォント
$fontFolder = "$env:SystemRoot\Fonts"
$fonts = Get-ChildItem $fontFolder | Select-Object Name, Extension, Length |
    ForEach-Object { [ordered]@{ Name=$_.Name; Extension=$_.Extension; SizeKB=[math]::Round($_.Length/1KB,1) } }
$f = Save-Json "17_fonts.json" $fonts
Append-Index "17. フォント" $f

# プリンター
$printers = Get-Printer -ErrorAction SilentlyContinue | ForEach-Object {
    [ordered]@{
        Name           = $_.Name
        DriverName     = $_.DriverName
        PortName       = $_.PortName
        Shared         = $_.Shared
        Default        = $_.Default
        PrinterStatus  = $_.PrinterStatus.ToString()
    }
}
$f = Save-Json "17b_printers.json" $printers
Append-Index "17b. プリンター" $f

# ─────────────────────────────────────────────
# 18. シェル・ターミナル設定
# ─────────────────────────────────────────────
Write-Step 18 "シェル・ターミナル設定"

# PowerShell プロファイル
$psProfiles = @()
$profilePaths = @(
    $PROFILE.AllUsersAllHosts,
    $PROFILE.AllUsersCurrentHost,
    $PROFILE.CurrentUserAllHosts,
    $PROFILE.CurrentUserCurrentHost
)
foreach ($profilePath in $profilePaths) {
    if (Test-Path $profilePath) {
        $psProfiles += "=== $profilePath ==="
        $psProfiles += (Get-Content $profilePath)
        $psProfiles += ""
    } else {
        $psProfiles += "=== $profilePath (存在しない) ==="
        $psProfiles += ""
    }
}
$f = Save-Text "18_powershell_profiles.txt" $psProfiles
Append-Index "18. PowerShell プロファイル" $f

# Windows Terminal 設定
$wtSettingsPaths = @(
    "$env:LOCALAPPDATA\Packages\Microsoft.WindowsTerminal_8wekyb3d8bbwe\LocalState\settings.json",
    "$env:LOCALAPPDATA\Packages\Microsoft.WindowsTerminalPreview_8wekyb3d8bbwe\LocalState\settings.json",
    "$env:APPDATA\Microsoft\Windows Terminal\settings.json"
)
foreach ($wtPath in $wtSettingsPaths) {
    if (Test-Path $wtPath) {
        Copy-Item $wtPath (Join-Path $outDir "18b_windows_terminal_settings.json")
        Append-Index "18b. Windows Terminal 設定" (Join-Path $outDir "18b_windows_terminal_settings.json")
        break
    }
}

# デフォルトブラウザ
$defaultBrowser = (Get-ItemProperty 'HKCU:\Software\Microsoft\Windows\Shell\Associations\UrlAssociations\https\UserChoice' -ErrorAction SilentlyContinue).ProgId
$shellSettings = [ordered]@{
    DefaultBrowserProgId = $defaultBrowser
    DefaultShell = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon' -ErrorAction SilentlyContinue).Shell
    FileExplorer = [ordered]@{
        ShowHiddenFiles = (Get-ItemProperty 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced' -ErrorAction SilentlyContinue).Hidden
        ShowFileExtensions = -not ((Get-ItemProperty 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced' -ErrorAction SilentlyContinue).HideFileExt)
        ShowFullPath = (Get-ItemProperty 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\CabinetState' -ErrorAction SilentlyContinue).FullPath
    }
}
$f = Save-Json "18c_shell_settings.json" $shellSettings
Append-Index "18c. シェル・エクスプローラー設定" $f

# 証明書（追加されたカスタム証明書）
$certStores = @('Cert:\LocalMachine\My', 'Cert:\LocalMachine\Root', 'Cert:\CurrentUser\My')
$certs = $certStores | ForEach-Object {
    $store = $_
    Get-ChildItem $store -ErrorAction SilentlyContinue | ForEach-Object {
        [ordered]@{
            Store          = $store
            Subject        = $_.Subject
            Thumbprint     = $_.Thumbprint
            NotBefore      = $_.NotBefore
            NotAfter       = $_.NotAfter
            Issuer         = $_.Issuer
            HasPrivateKey  = $_.HasPrivateKey
        }
    }
}
$f = Save-Json "18d_certificates.json" $certs
Append-Index "18d. 証明書" $f

# ─────────────────────────────────────────────
# 19. ポリシー設定
# ─────────────────────────────────────────────
Write-Step 19 "ポリシー設定"

# gpresult テキスト形式（ドメイン参加・非参加どちらでも動作）
$gpresultLines = @("=== gpresult /r ===")
$gpresultLines += (gpresult /r 2>&1)
$f = Save-Text "19_gpresult.txt" $gpresultLines
Append-Index "19. GPO 適用結果（テキスト）" $f

# gpresult HTML形式（詳細レポート）
$gpresultHtml = Join-Path $outDir "19b_gpresult.html"
gpresult /h $gpresultHtml /f 2>&1 | Out-Null
if (Test-Path $gpresultHtml) {
    Append-Index "19b. GPO 適用結果（HTML詳細）" $gpresultHtml
}

# レジストリポリシーキースキャン（HKLM/HKCU の Policies キー配下）
function Get-RegistryPolicies {
    param([string]$rootPath)
    if (-not (Test-Path $rootPath)) { return @() }
    $results = @()
    Get-ChildItem $rootPath -Recurse -ErrorAction SilentlyContinue | ForEach-Object {
        $keyPath = $_.PSPath -replace 'Microsoft\.PowerShell\.Core\\Registry::', ''
        $props   = Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue
        if ($props) {
            $props.PSObject.Properties |
                Where-Object { $_.Name -notmatch '^PS' } |
                ForEach-Object {
                    $results += [ordered]@{
                        KeyPath = $keyPath
                        Name    = $_.Name
                        Value   = "$($_.Value)"
                    }
                }
        }
    }
    return $results
}

$policyRoots = @(
    'HKLM:\SOFTWARE\Policies',
    'HKCU:\SOFTWARE\Policies',
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies',
    'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies'
)
$allPolicies = [ordered]@{
    ComputerPolicies = @(Get-RegistryPolicies $policyRoots[0])
    UserPolicies     = @(Get-RegistryPolicies $policyRoots[1])
    ComputerPoliciesLegacy = @(Get-RegistryPolicies $policyRoots[2])
    UserPoliciesLegacy     = @(Get-RegistryPolicies $policyRoots[3])
}
$f = Save-Json "19c_registry_policies.json" $allPolicies
Append-Index "19c. レジストリポリシー (SOFTWARE\Policies)" $f

# 監査ポリシー（管理者権限不要）
$auditLines = @("=== auditpol /get /category:* ===")
$auditLines += (auditpol /get /category:* 2>&1)
$f = Save-Text "19d_audit_policy.txt" $auditLines
Append-Index "19d. 監査ポリシー (auditpol)" $f

# パスワード・アカウントロックアウトポリシー（管理者権限不要）
$netAccLines = @("=== net accounts ===")
$netAccLines += (net accounts 2>&1)
$f = Save-Text "19e_account_policy.txt" $netAccLines
Append-Index "19e. アカウントポリシー (net accounts)" $f

# ローカルセキュリティポリシー全体（管理者権限必要）
if ($isAdmin) {
    $seceditInf = Join-Path $outDir "19f_local_security_policy.inf"
    secedit /export /cfg $seceditInf /quiet 2>&1 | Out-Null
    if (Test-Path $seceditInf) {
        Append-Index "19f. ローカルセキュリティポリシー (secedit)" $seceditInf
    }
}

# ─────────────────────────────────────────────
# 20. スクリーンショット
# ─────────────────────────────────────────────
Write-Step 20 "スクリーンショット"

Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue
Add-Type -AssemblyName System.Drawing        -ErrorAction SilentlyContinue
Add-Type @"
using System;
using System.Runtime.InteropServices;
public class ScreenCapHelper {
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern bool SetProcessDPIAware();
    [DllImport("user32.dll")] public static extern int GetSystemMetrics(int nIndex);
}
"@ -ErrorAction SilentlyContinue

function Save-Screenshot {
    param([string]$filepath)
    [ScreenCapHelper]::SetProcessDPIAware() | Out-Null
    $width  = [ScreenCapHelper]::GetSystemMetrics(0)  # SM_CXSCREEN (物理ピクセル)
    $height = [ScreenCapHelper]::GetSystemMetrics(1)  # SM_CYSCREEN (物理ピクセル)
    $bitmap = New-Object System.Drawing.Bitmap($width, $height)
    $g      = [System.Drawing.Graphics]::FromImage($bitmap)
    $g.CopyFromScreen(0, 0, 0, 0, [System.Drawing.Size]::new($width, $height))
    $bitmap.Save($filepath, [System.Drawing.Imaging.ImageFormat]::Png)
    $g.Dispose(); $bitmap.Dispose()
}

# スタートメニューを開いた状態のスクリーンショット
[System.Windows.Forms.SendKeys]::SendWait("^{ESC}")
Start-Sleep -Milliseconds 1500
$ssStart = Join-Path $outDir "20_screenshot_startmenu.png"
Save-Screenshot $ssStart
Append-Index "20. スクリーンショット（スタートメニュー）" $ssStart
[System.Windows.Forms.SendKeys]::SendWait("{ESC}")
Start-Sleep -Milliseconds 500

# 設定アプリを全画面で開いた状態のスクリーンショット
Start-Process "ms-settings:"
Start-Sleep -Seconds 2
$settingsProc = Get-Process "SystemSettings" -ErrorAction SilentlyContinue | Select-Object -First 1
if ($settingsProc -and $settingsProc.MainWindowHandle -ne 0) {
    [ScreenCapHelper]::SetForegroundWindow($settingsProc.MainWindowHandle) | Out-Null
    [ScreenCapHelper]::ShowWindow($settingsProc.MainWindowHandle, 3) | Out-Null  # SW_MAXIMIZE
    Start-Sleep -Milliseconds 800
}
$ssSettings = Join-Path $outDir "20b_screenshot_settings.png"
Save-Screenshot $ssSettings
Append-Index "20b. スクリーンショット（設定）" $ssSettings
if ($settingsProc) { $settingsProc | Stop-Process -Force -ErrorAction SilentlyContinue }

# ─────────────────────────────────────────────
# インデックスファイル生成
# ─────────────────────────────────────────────
Write-Progress -Activity "環境調査中" -Completed
$indexLines += ""
$indexLines += "---"
$indexLines += "収集完了: $(Get-Date)"
$indexPath = Join-Path $outDir "00_INDEX.md"
$indexLines | Set-Content -Path $indexPath -Encoding UTF8

# ─────────────────────────────────────────────
# サマリー表示
# ─────────────────────────────────────────────
$files = Get-ChildItem $outDir | Measure-Object
$sizeMB = [math]::Round(((Get-ChildItem $outDir | Measure-Object Length -Sum).Sum) / 1MB, 2)

Write-Host ""
Write-Host "=" * 60 -ForegroundColor Green
Write-Host " 調査完了" -ForegroundColor Green
Write-Host "=" * 60 -ForegroundColor Green
Write-Host " 出力先  : $outDir"
Write-Host " ファイル数: $($files.Count) 個"
Write-Host " 合計サイズ: $sizeMB MB"
if (-not $isAdmin) {
    Write-Host ""
    Write-Host " ※ 管理者権限なしで実行されたため、以下は未収集です:" -ForegroundColor Yellow
    Write-Host "   - BitLocker 状態"
    Write-Host "   - ファイアウォールカスタムルール"
    Write-Host "   - 一部の WMI 情報"
    Write-Host "   管理者権限で再実行すると完全な情報を収集できます。" -ForegroundColor Yellow
}
Write-Host ""
Write-Host " インデックス: $indexPath"
Write-Host "=" * 60 -ForegroundColor Green
