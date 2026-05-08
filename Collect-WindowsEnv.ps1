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

.PARAMETER GenerateHtml
    指定すると収集結果をまとめた HTML レポート（00_REPORT.html）を追加生成する

.EXAMPLE
    .\Collect-WindowsEnv.ps1
    .\Collect-WindowsEnv.ps1 -OutputPath "D:\EnvBackup"
    .\Collect-WindowsEnv.ps1 -GenerateHtml
    .\Collect-WindowsEnv.ps1 -OutputPath "D:\EnvBackup" -GenerateHtml
#>
param(
    [string]$OutputPath = $PSScriptRoot,
    [switch]$GenerateHtml
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
$total      = 22

function Write-Step {
    param([int]$n, [string]$name)
    $script:progress = $n
    Write-Progress -Activity "環境調査中" -Status "[$n/$total] $name" -PercentComplete ([int]($n / $total * 100))
    Write-Host "[$n/$total] $name ..." -ForegroundColor Cyan
}

function Save-Json {
    param([string]$filename, $data, [string]$emptyReason = $null)
    $path = Join-Path $outDir $filename

    # 空・null 時に出力する診断オブジェクトを生成する内部関数
    function Get-DiagJson {
        param([string]$reason)
        $r = if ($reason) { $reason } else {
            if (-not $isAdmin) {
                'データなし。管理者権限なしで実行されたため取得できなかった可能性があります。管理者権限（Run as Administrator）で再実行してください。'
            } else {
                'データなし。この環境には該当する項目が存在しないか、コマンドが失敗した可能性があります。'
            }
        }
        [ordered]@{
            _status    = 'データなし'
            _reason    = $r
            _timestamp = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        } | ConvertTo-Json -Depth 3
    }

    # @() | ConvertTo-Json はパイプラインが空になりファイルが作成されないため明示的に処理
    $json = if ($null -eq $data) {
        Get-DiagJson $emptyReason
    } elseif ($data -is [System.Collections.IEnumerable] -and -not ($data -is [string])) {
        $arr = @($data)
        if ($arr.Count -eq 0) { Get-DiagJson $emptyReason } else { $arr | ConvertTo-Json -Depth 10 }
    } else {
        $data | ConvertTo-Json -Depth 10
    }
    if ($null -eq $json) { $json = Get-DiagJson $emptyReason }
    $json | Set-Content -Path $path -Encoding UTF8
    return $path
}

function Save-Text {
    param([string]$filename, [string[]]$lines, [string]$emptyReason = $null)
    $path = Join-Path $outDir $filename
    # ヘッダー行（=== ... ===）と空行のみで実データがない場合は診断情報を追記する
    $dataLines = @($lines | Where-Object { $_ -and $_.Trim() -ne '' -and $_ -notmatch '^===' -and $_ -notmatch '^---' })
    if ($dataLines.Count -eq 0) {
        $reason = if ($emptyReason) { $emptyReason } else {
            if (-not $isAdmin) {
                'コマンドが出力を返しませんでした。管理者権限（Run as Administrator）で再実行すると取得できる可能性があります。'
            } else {
                'コマンドが出力を返しませんでした。この環境では該当する設定が存在しないか、コマンドが利用できない可能性があります。'
            }
        }
        $lines = @($lines) + @('', '=== データなし ===', "理由: $reason", "記録日時: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
    }
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

$disks = try {
    Get-Disk -ErrorAction Stop | ForEach-Object {
        $disk = $_
        $parts = Get-Partition -DiskNumber $disk.Number -ErrorAction SilentlyContinue
        [ordered]@{
            DiskNumber       = $disk.Number
            FriendlyName     = $disk.FriendlyName
            SizeGB           = [math]::Round($disk.Size / 1GB, 2)
            PartitionStyle   = $disk.PartitionStyle
            BusType          = $disk.BusType
            MediaType        = if ($disk.PSObject.Properties['MediaType']) { $disk.MediaType } else { $null }
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
                    FileSystem      = if ($vol -and $vol.PSObject.Properties['FileSystem']) { $vol.FileSystem } else { $null }
                    Label           = if ($vol -and $vol.PSObject.Properties['FileSystemLabel']) { $vol.FileSystemLabel } else { $null }
                }
            })
        }
    }
} catch {
    # Get-Disk 失敗時は WMI（Win32_DiskDrive）でフォールバック（パーティション情報も取得）
    Write-Warning "[02] Get-Disk 失敗。WMI から取得します: $_"
    Get-CimInstance Win32_DiskDrive -ErrorAction SilentlyContinue | ForEach-Object {
        $drive = $_
        # Win32_DiskDriveToDiskPartition → Win32_LogicalDiskToPartition で論理ドライブまで辿る
        $partitions = Get-CimAssociatedInstance -InputObject $drive -ResultClassName Win32_DiskPartition `
            -ErrorAction SilentlyContinue | ForEach-Object {
            $part = $_
            $logicals = Get-CimAssociatedInstance -InputObject $part -ResultClassName Win32_LogicalDisk `
                -ErrorAction SilentlyContinue
            $logical1 = if ($logicals) { @($logicals)[0] } else { $null }
            [ordered]@{
                PartitionNumber = $part.Index
                Type            = $part.Type
                SizeGB          = [math]::Round($part.Size / 1GB, 2)
                DriveLetter     = if ($logical1 -and $logical1.PSObject.Properties['DeviceID']) { $logical1.DeviceID -replace ':','' } else { $null }
                IsSystem        = $part.BootPartition
                IsBoot          = $part.PrimaryPartition
                FileSystem      = if ($logical1 -and $logical1.PSObject.Properties['FileSystem']) { $logical1.FileSystem } else { $null }
                Label           = if ($logical1 -and $logical1.PSObject.Properties['VolumeName']) { $logical1.VolumeName } else { $null }
                FreeGB          = if ($logical1 -and $logical1.PSObject.Properties['FreeSpace'] -and $logical1.FreeSpace) { [math]::Round($logical1.FreeSpace / 1GB, 2) } else { $null }
            }
        }
        [ordered]@{
            DiskNumber       = $drive.Index
            FriendlyName     = $drive.Model
            SizeGB           = if ($drive.Size) { [math]::Round($drive.Size / 1GB, 2) } else { 0 }
            PartitionStyle   = $null
            BusType          = $drive.InterfaceType
            MediaType        = if ($drive.PSObject.Properties['MediaType'])  { $drive.MediaType }  else { $null }
            SerialNumber     = if ($drive.PSObject.Properties['SerialNumber']){ $drive.SerialNumber } else { $null }
            OperationalStatus= $drive.Status
            Partitions       = @($partitions)
        }
    }
}
$f = Save-Json "02_disk_partitions.json" $disks
Append-Index "02. ディスク・パーティション" $f

# ─────────────────────────────────────────────
# 02b. 論理ドライブ一覧（全ドライブレター・空き容量）
# ─────────────────────────────────────────────
$logicalDisks = Get-CimInstance Win32_LogicalDisk -ErrorAction SilentlyContinue | ForEach-Object {
    [ordered]@{
        DeviceID    = $_.DeviceID
        DriveType   = switch ($_.DriveType) {
            0 { 'Unknown' } 1 { 'NoRootDir' } 2 { 'Removable' }
            3 { 'LocalDisk' } 4 { 'Network' } 5 { 'CDRom' } 6 { 'RAMDisk' }
            default { "$($_.DriveType)" }
        }
        FileSystem  = $_.FileSystem
        VolumeName  = $_.VolumeName
        SizeGB      = if ($_.Size) { [math]::Round($_.Size / 1GB, 2) } else { 0 }
        FreeGB      = if ($_.FreeSpace) { [math]::Round($_.FreeSpace / 1GB, 2) } else { 0 }
        UsedPercent = if ($_.Size -and $_.Size -gt 0) { [math]::Round((($_.Size - $_.FreeSpace) / $_.Size) * 100, 1) } else { $null }
        VolumeSerialNumber = $_.VolumeSerialNumber
    }
}
$f = Save-Json "02b_logical_disks.json" $logicalDisks
Append-Index "02b. 論理ドライブ一覧" $f

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

$storeAppsError = $null
$storeApps = try {
    Get-AppxPackage -AllUsers -ErrorAction Stop |
        Select-Object Name, Version, Publisher, Architecture, PackageFullName |
        Sort-Object Name | ForEach-Object {
            [ordered]@{
                Name            = $_.Name
                Version         = $_.Version
                Publisher       = $_.Publisher
                Architecture    = $_.Architecture
                PackageFullName = $_.PackageFullName
            }
        }
} catch { $storeAppsError = $_.Exception.Message; @() }
$storeAppsReason = if (@($storeApps).Count -eq 0) {
    $r = 'Store アプリ（AppX/MSIX）一覧を取得できませんでした'
    if ($storeAppsError) { $r += "。エラー: $storeAppsError" }
    elseif (-not $isAdmin) { $r += '。-AllUsers オプションには管理者権限が必要です' }
    $r
} else { $null }
$f = Save-Json "04_installed_apps_store.json" $storeApps $storeAppsReason
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
    $wingetRaw = (winget list --accept-source-agreements 2>&1)
    # プログレスバー表示行（█▒░ 等のブロック文字を含む行）を除去してから記録
    $pkgLines += ($wingetRaw | Where-Object { $_ -notmatch '[█▒░]' })
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

$optFeatError = $null
$optFeatures = try {
    Get-WindowsOptionalFeature -Online -ErrorAction Stop |
        Select-Object FeatureName, State | Sort-Object FeatureName |
        ForEach-Object { [ordered]@{ Feature=$_.FeatureName; State=$_.State.ToString() } }
} catch { $optFeatError = $_.Exception.Message; @() }
$optFeatReason = if (@($optFeatures).Count -eq 0) {
    $r = 'Windows オプション機能の一覧を取得できませんでした'
    if ($optFeatError) { $r += "。エラー: $optFeatError" }
    elseif (-not $isAdmin) { $r += '。管理者権限が必要です' }
    $r
} else { $null }
$f = Save-Json "07_windows_optional_features.json" $optFeatures $optFeatReason
Append-Index "07. Windows オプション機能" $f

# Windows Capability
$capabilityError = $null
$capabilities = try {
    Get-WindowsCapability -Online -ErrorAction Stop |
        Select-Object Name, State | Sort-Object Name |
        ForEach-Object { [ordered]@{ Name=$_.Name; State=$_.State.ToString() } }
} catch { $capabilityError = $_.Exception.Message; @() }
$capabilityReason = if (@($capabilities).Count -eq 0) {
    $r = 'Windows Capability の一覧を取得できませんでした'
    if ($capabilityError) { $r += "。エラー: $capabilityError" }
    elseif (-not $isAdmin) { $r += '。管理者権限が必要です' }
    $r
} else { $null }
$f = Save-Json "07b_windows_capabilities.json" $capabilities $capabilityReason
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
$netAdapters = try {
    Get-NetAdapter -ErrorAction Stop | Where-Object { $_.Status -ne 'Not Present' } | ForEach-Object {
        $adapter = $_
        $ipConfig = Get-NetIPConfiguration -InterfaceIndex $adapter.InterfaceIndex -ErrorAction SilentlyContinue
        $dnsServers = (Get-DnsClientServerAddress -InterfaceIndex $adapter.InterfaceIndex -AddressFamily IPv4 -ErrorAction SilentlyContinue).ServerAddresses
        $linkSpeedMbps = try {
            $ls = $adapter.LinkSpeed
            if ($null -eq $ls) { 0 }
            elseif ($ls -is [long] -or $ls -is [int] -or $ls -is [double]) {
                if ($ls -gt 0) { [math]::Round($ls / 1MB, 0) } else { 0 }
            } else { "$ls" }
        } catch { if ($adapter.LinkSpeed) { "$($adapter.LinkSpeed)" } else { 0 } }
        [ordered]@{
            Name             = $adapter.Name
            InterfaceAlias   = $adapter.InterfaceAlias
            InterfaceDescription = $adapter.InterfaceDescription
            Status           = $adapter.Status.ToString()
            MacAddress       = $adapter.MacAddress
            LinkSpeedMbps    = $linkSpeedMbps
            IPv4Address      = @(if ($ipConfig -and $ipConfig.IPv4Address)        { $ipConfig.IPv4Address.IPAddress })
            IPv4PrefixLength = @(if ($ipConfig -and $ipConfig.IPv4Address)        { $ipConfig.IPv4Address.PrefixLength })
            DefaultGateway   = @(if ($ipConfig -and $ipConfig.IPv4DefaultGateway) { $ipConfig.IPv4DefaultGateway.NextHop })
            DNSServers       = @(if ($dnsServers) { $dnsServers })
            IPv6Address      = @(if ($ipConfig -and $ipConfig.IPv6Address)        { $ipConfig.IPv6Address.IPAddress })
        }
    }
} catch {
    Write-Warning "[09] Get-NetAdapter 失敗: $_"
    @()
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
    $fwRulesReason = if (@($fwRules).Count -eq 0) {
        'カスタムファイアウォールルールが存在しないか、Get-NetFirewallRule の実行に失敗しました。' +
        'すべてのルールが組み込み（ビルトイン）であり、ローカルに追加されたルールがない可能性があります。'
    } else { $null }
    $f = Save-Json "09e_firewall_rules_custom.json" $fwRules $fwRulesReason
    Append-Index "09e. ファイアウォールカスタムルール" $f
}

# ─────────────────────────────────────────────
# 09f. ファイアウォールプロファイル（各プロファイルの有効/無効）
# ─────────────────────────────────────────────
$fwProfiles = try {
    Get-NetFirewallProfile -ErrorAction Stop | ForEach-Object {
        [ordered]@{
            Name                  = $_.Name
            Enabled               = $_.Enabled.ToString()
            DefaultInboundAction  = $_.DefaultInboundAction.ToString()
            DefaultOutboundAction = $_.DefaultOutboundAction.ToString()
            LogAllowed            = $_.LogAllowed.ToString()
            LogBlocked            = $_.LogBlocked.ToString()
            LogFileName           = $_.LogFileName
            LogMaxSizeKilobytes   = $_.LogMaxSizeKilobytes
        }
    }
} catch { @() }
$f = Save-Json "09f_firewall_profiles.json" $fwProfiles
Append-Index "09f. ファイアウォールプロファイル" $f

# ─────────────────────────────────────────────
# 09g. 登録済みセキュリティ製品 (SecurityCenter2 WMI)
#      Windows セキュリティセンターに登録されたAV・FW・AS製品一覧
# ─────────────────────────────────────────────
function ConvertTo-SecurityProductState {
    # productState の16進数フィールドを解読する
    # 上位バイト: 0x10=有効, 0x11=無効
    # 中位バイト: 0x00=最新, 0x10=期限切れ
    param([int]$state)
    $enabled  = (($state -band 0x1000) -ne 0)
    $upToDate = (($state -band 0x0010) -eq 0)
    "$state (有効:$enabled / 定義最新:$upToDate)"
}
$secProducts = [ordered]@{
    AntiVirus = @(
        Get-CimInstance -Namespace 'root/SecurityCenter2' -ClassName AntiVirusProduct -ErrorAction SilentlyContinue |
        ForEach-Object {
            [ordered]@{
                Name         = $_.displayName
                InstanceGuid = $_.instanceGuid
                ExePath      = $_.pathToSignedProductExe
                ProductState = ConvertTo-SecurityProductState ([int]$_.productState)
            }
        }
    )
    Firewall = @(
        Get-CimInstance -Namespace 'root/SecurityCenter2' -ClassName FirewallProduct -ErrorAction SilentlyContinue |
        ForEach-Object {
            [ordered]@{
                Name         = $_.displayName
                InstanceGuid = $_.instanceGuid
                ExePath      = $_.pathToSignedProductExe
                ProductState = ConvertTo-SecurityProductState ([int]$_.productState)
            }
        }
    )
    AntiSpyware = @(
        Get-CimInstance -Namespace 'root/SecurityCenter2' -ClassName AntiSpywareProduct -ErrorAction SilentlyContinue |
        ForEach-Object {
            [ordered]@{
                Name         = $_.displayName
                InstanceGuid = $_.instanceGuid
                ExePath      = $_.pathToSignedProductExe
                ProductState = ConvertTo-SecurityProductState ([int]$_.productState)
            }
        }
    )
}
$f = Save-Json "09g_security_products.json" $secProducts
Append-Index "09g. 登録済みセキュリティ製品 (SecurityCenter2)" $f

# ─────────────────────────────────────────────
# 09h. Windows Defender ステータス (Get-MpComputerStatus)
# ─────────────────────────────────────────────
$mpStatus = try { Get-MpComputerStatus -ErrorAction Stop } catch { $null }
$defenderStatus = if ($mpStatus) {
    [ordered]@{
        # 保護機能のON/OFF
        AMServiceEnabled          = $mpStatus.AMServiceEnabled
        AntivirusEnabled          = $mpStatus.AntivirusEnabled
        AntispywareEnabled        = $mpStatus.AntispywareEnabled
        RealTimeProtectionEnabled = $mpStatus.RealTimeProtectionEnabled
        BehaviorMonitorEnabled    = $mpStatus.BehaviorMonitorEnabled
        IoavProtectionEnabled     = $mpStatus.IoavProtectionEnabled
        NISEnabled                = $mpStatus.NISEnabled
        OnAccessProtectionEnabled = $mpStatus.OnAccessProtectionEnabled
        TamperProtectionSource    = "$($mpStatus.TamperProtectionSource)"
        # バージョン・署名
        AMProductVersion              = $mpStatus.AMProductVersion
        AMEngineVersion               = $mpStatus.AMEngineVersion
        AMServiceVersion              = $mpStatus.AMServiceVersion
        AntivirusSignatureVersion     = $mpStatus.AntivirusSignatureVersion
        AntispywareSignatureVersion   = $mpStatus.AntispywareSignatureVersion
        NISSignatureVersion           = $mpStatus.NISSignatureVersion
        AntivirusSignatureLastUpdated = "$($mpStatus.AntivirusSignatureLastUpdated)"
        # スキャン履歴
        FullScanEndTime    = "$($mpStatus.FullScanEndTime)"
        QuickScanEndTime   = "$($mpStatus.QuickScanEndTime)"
        # 脅威・検疫
        QuarantineCount             = $mpStatus.QuarantineCount
        ThreatCount                 = $mpStatus.ThreatCount
        DefenderSignaturesOutOfDate = $mpStatus.DefenderSignaturesOutOfDate
        FullScanRequired            = $mpStatus.FullScanRequired
        RebootRequired              = $mpStatus.RebootRequired
        ComputerState               = $mpStatus.ComputerState
    }
} else {
    [ordered]@{ Error = "Get-MpComputerStatus 失敗（Defender 未インストールまたは権限不足）" }
}
$f = Save-Json "09h_defender_status.json" $defenderStatus
Append-Index "09h. Windows Defender ステータス" $f

# ─────────────────────────────────────────────
# 09i. Windows Defender 設定 (Get-MpPreference)
#      除外設定・ASRルール・クラウド保護・ネットワーク保護 等
# ─────────────────────────────────────────────
$mpPref = try { Get-MpPreference -ErrorAction Stop } catch { $null }
$defenderPref = if ($mpPref) {
    [ordered]@{
        # 除外設定（組織管理で重要）
        ExclusionPath        = @($mpPref.ExclusionPath)
        ExclusionExtension   = @($mpPref.ExclusionExtension)
        ExclusionProcess     = @($mpPref.ExclusionProcess)
        ExclusionIpAddress   = @($mpPref.ExclusionIpAddress)
        # リアルタイム保護
        DisableRealtimeMonitoring   = $mpPref.DisableRealtimeMonitoring
        DisableBehaviorMonitoring   = $mpPref.DisableBehaviorMonitoring
        DisableBlockAtFirstSeen     = $mpPref.DisableBlockAtFirstSeen
        DisableIOAVProtection       = $mpPref.DisableIOAVProtection
        DisableScriptScanning       = $mpPref.DisableScriptScanning
        # クラウド保護 (MAPS)
        MAPSReporting        = "$($mpPref.MAPSReporting)"
        CloudBlockLevel      = "$($mpPref.CloudBlockLevel)"
        CloudExtendedTimeout = $mpPref.CloudExtendedTimeout
        SubmitSamplesConsent = "$($mpPref.SubmitSamplesConsent)"
        # ネットワーク保護
        EnableNetworkProtection = "$($mpPref.EnableNetworkProtection)"
        # PUA 保護
        PUAProtection        = "$($mpPref.PUAProtection)"
        # Attack Surface Reduction (ASR) ルール
        ASR_RuleIds    = @($mpPref.AttackSurfaceReductionRules_Ids)
        ASR_Actions    = @($mpPref.AttackSurfaceReductionRules_Actions)
        # スキャンスケジュール
        ScanScheduleDay  = "$($mpPref.ScanScheduleDay)"
        ScanScheduleTime = "$($mpPref.ScanScheduleTime)"
        SignatureScheduleDay = "$($mpPref.SignatureScheduleDay)"
    }
} else {
    [ordered]@{ Error = "Get-MpPreference 失敗（Defender 未インストールまたは権限不足）" }
}
$f = Save-Json "09i_defender_preferences.json" $defenderPref
Append-Index "09i. Windows Defender 設定 (Get-MpPreference)" $f

# ─────────────────────────────────────────────
# 09j. Microsoft Defender for Endpoint (MDE) オンボーディング状態
# ─────────────────────────────────────────────
$mdeStatusReg  = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows Advanced Threat Protection\Status' -ErrorAction SilentlyContinue
$mdeOnboarding = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows Advanced Threat Protection' -ErrorAction SilentlyContinue

$senseService  = Get-Service -Name 'Sense'   -ErrorAction SilentlyContinue
$msSenseService= Get-Service -Name 'MsSense' -ErrorAction SilentlyContinue

$mdeInfo = [ordered]@{
    # サービス状態（Sense = MDE センサー）
    SenseService   = if ($senseService)   { [ordered]@{ Status=$senseService.Status.ToString();   StartType=$senseService.StartType.ToString() } }   else { $null }
    MsSenseService = if ($msSenseService) { [ordered]@{ Status=$msSenseService.Status.ToString(); StartType=$msSenseService.StartType.ToString() } } else { $null }
    # オンボーディング状態（0=未オンボーディング, 1=オンボーディング済み）
    OnboardingState = $mdeStatusReg.OnboardingState
    OrgId           = $mdeStatusReg.OrgId
    SenseId         = $mdeStatusReg.SenseId
    # 最終ハートビート
    LastConnected   = "$($mdeStatusReg.LastConnected)"
}
$f = Save-Json "09j_defender_endpoint.json" $mdeInfo
Append-Index "09j. Microsoft Defender for Endpoint (MDE) 状態" $f

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
    $bitlockerReason = if (-not $bitlocker -or @($bitlocker).Count -eq 0) {
        'BitLocker が有効なボリュームが存在しないか、Get-BitLockerVolume の実行に失敗しました。' +
        'BitLocker が無効（すべてのドライブで未暗号化）の場合も空になります。'
    } else { $null }
    $f = Save-Json "10e_bitlocker.json" $bitlocker $bitlockerReason
    Append-Index "10e. BitLocker" $f
}

# ─────────────────────────────────────────────
# 10f. 全ユーザープロファイル一覧
# ─────────────────────────────────────────────
$profileListKey = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList'
$allProfiles = Get-CimInstance Win32_UserProfile -ErrorAction SilentlyContinue | ForEach-Object {
    $prof   = $_
    $sid    = $prof.SID
    $regKey = Get-ItemProperty "$profileListKey\$sid" -ErrorAction SilentlyContinue
    # mandatory プロファイルはパスが .man で終わるか、CentralProfile が .man
    $isMandatory = ($prof.LocalPath -match '\.man$') -or ($regKey.CentralProfile -match '\.man$')
    [ordered]@{
        LocalPath        = $prof.LocalPath
        SID              = $sid
        Loaded           = $prof.Loaded
        RoamingConfigured= $prof.RoamingConfigured
        Special          = $prof.Special
        LastUseTime      = $prof.LastUseTime
        MandatoryProfile = $isMandatory
        CentralProfile   = $regKey.CentralProfile
    }
}
$f = Save-Json "10f_user_profiles.json" $allProfiles
Append-Index "10f. 全ユーザープロファイル一覧" $f

# ─────────────────────────────────────────────
# 10g. 自動ログイン設定
# ─────────────────────────────────────────────
$winlogonKey = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon' -ErrorAction SilentlyContinue
$autoLogin = [ordered]@{
    AutoAdminLogon    = $winlogonKey.AutoAdminLogon
    DefaultUserName   = $winlogonKey.DefaultUserName
    DefaultDomainName = $winlogonKey.DefaultDomainName
    # パスワードは存在有無のみ記録（平文保存のため値は収集しない）
    DefaultPassword   = if ($winlogonKey.DefaultPassword) { '*** (設定あり)' } else { $null }
    ForceAutoLogon    = $winlogonKey.ForceAutoLogon
    Userinit          = $winlogonKey.Userinit
    Shell             = $winlogonKey.Shell
}
$f = Save-Json "10g_autologin_settings.json" $autoLogin
Append-Index "10g. 自動ログイン設定 (Winlogon)" $f

# ─────────────────────────────────────────────
# 10h. ユーザーごとのデスクトップ・スタートアップ・Run レジストリ
# ─────────────────────────────────────────────
$perUserDetails = @()
$allProfiles | Where-Object { $_.LocalPath -and (Test-Path $_.LocalPath) -and -not $_.Special } | ForEach-Object {
    $userPath = $_.LocalPath
    $userName = Split-Path $userPath -Leaf
    $sid      = $_.SID

    # デスクトップ上のファイル・ショートカット一覧
    $desktopFiles = @()
    $desktopPath = Join-Path $userPath 'Desktop'
    if (Test-Path $desktopPath) {
        $shell = New-Object -ComObject WScript.Shell -ErrorAction SilentlyContinue
        $desktopFiles = Get-ChildItem $desktopPath -Recurse -Depth 2 -ErrorAction SilentlyContinue | ForEach-Object {
            $lnkTarget = $null
            if ($_.Extension -eq '.lnk' -and $shell) {
                $lnk = Try-Command { $shell.CreateShortcut($_.FullName) } $null
                $lnkTarget = if ($lnk) { $lnk.TargetPath } else { $null }
            }
            [ordered]@{
                Name           = $_.Name
                FullPath       = $_.FullName
                Extension      = $_.Extension
                LastModified   = $_.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss')
                ShortcutTarget = $lnkTarget
            }
        }
    }

    # ユーザースタートアップフォルダ
    $userStartupFiles = @()
    $userStartupPath = Join-Path $userPath 'AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup'
    if (Test-Path $userStartupPath) {
        $shell2 = New-Object -ComObject WScript.Shell -ErrorAction SilentlyContinue
        $userStartupFiles = Get-ChildItem $userStartupPath -ErrorAction SilentlyContinue | ForEach-Object {
            $lnkTarget2 = $null
            if ($_.Extension -eq '.lnk' -and $shell2) {
                $lnk2 = Try-Command { $shell2.CreateShortcut($_.FullName) } $null
                $lnkTarget2 = if ($lnk2) { $lnk2.TargetPath } else { $null }
            }
            [ordered]@{
                Name           = $_.Name
                FullPath       = $_.FullName
                ShortcutTarget = $lnkTarget2
            }
        }
    }

    # HKU\<SID> の Run / RunOnce（マウント済みハイブのみ取得可能）
    $hkuRun = @()
    foreach ($runKey in @("Run", "RunOnce")) {
        $keyPath = "Registry::HKU\$sid\SOFTWARE\Microsoft\Windows\CurrentVersion\$runKey"
        $props = Get-ItemProperty $keyPath -ErrorAction SilentlyContinue
        if ($props) {
            $props.PSObject.Properties | Where-Object { $_.Name -notmatch '^PS' } | ForEach-Object {
                $hkuRun += [ordered]@{ RegistryKey=$runKey; Name=$_.Name; Command=$_.Value }
            }
        }
    }

    $perUserDetails += [ordered]@{
        UserName         = $userName
        ProfilePath      = $userPath
        SID              = $sid
        Loaded           = $_.Loaded
        MandatoryProfile = $_.MandatoryProfile
        DesktopFiles     = @($desktopFiles)
        StartupFiles     = @($userStartupFiles)
        StartupRegistry  = @($hkuRun)
    }
}
$f = Save-Json "10h_per_user_details.json" $perUserDetails
Append-Index "10h. ユーザーごとの詳細（デスクトップ・スタートアップ・Run レジストリ）" $f

# ─────────────────────────────────────────────
# 10i. GPO ログオン・ログオフ・スタートアップ・シャットダウンスクリプト
# ─────────────────────────────────────────────
$gpoScripts = @()
$gpoMissingCategories = @()
foreach ($category in @('Logon','Logoff','Startup','Shutdown')) {
    $gpoPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\Scripts\$category"
    if (-not (Test-Path $gpoPath)) { $gpoMissingCategories += $category; continue }
    Get-ChildItem $gpoPath -ErrorAction SilentlyContinue | ForEach-Object {
        $gpoEntry = $_
        Get-ChildItem $gpoEntry.PSPath -ErrorAction SilentlyContinue | ForEach-Object {
            $scriptEntry = Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue
            $gpoScripts += [ordered]@{
                Category   = $category
                GPO        = $gpoEntry.PSChildName
                Script     = $scriptEntry.Script
                Parameters = $scriptEntry.Parameters
            }
        }
    }
}
$gpoScriptsReason = if ($gpoScripts.Count -eq 0 -and $gpoMissingCategories.Count -gt 0) {
    "レジストリキーが存在しないカテゴリ: $($gpoMissingCategories -join ', ')。" +
    'このキーはグループポリシークライアントがポリシー適用時に書き込みます。' +
    'ローカルグループポリシー（gpedit.msc）で設定されたスクリプトは再起動後に反映されます。' +
    'GPResult（19b_gpresult.html）に記録されていてもこのキーが空の場合は、' +
    'ポリシーがまだ適用されていないか、スクリプトの実体は C:\Windows\System32\GroupPolicy\Machine\Scripts\ に格納されています。'
} else { $null }
$f = Save-Json "10i_gpo_logon_scripts.json" $gpoScripts $gpoScriptsReason
Append-Index "10i. GPO ログオン・ログオフスクリプト" $f

# ─────────────────────────────────────────────
# 10i2. ユーザー側 GPO ログオン・ログオフスクリプト（HKCU）
# ─────────────────────────────────────────────
$userGpoScripts = @()
$userGpoMissingCategories = @()
foreach ($category in @('Logon','Logoff')) {
    $gpoPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Group Policy\Scripts\$category"
    if (-not (Test-Path $gpoPath)) { $userGpoMissingCategories += $category; continue }
    Get-ChildItem $gpoPath -ErrorAction SilentlyContinue | ForEach-Object {
        $gpoEntry = $_
        Get-ChildItem $gpoEntry.PSPath -ErrorAction SilentlyContinue | ForEach-Object {
            $scriptEntry = Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue
            $userGpoScripts += [ordered]@{
                Category   = $category
                GPO        = $gpoEntry.PSChildName
                Script     = $scriptEntry.Script
                Parameters = $scriptEntry.Parameters
            }
        }
    }
}
$userGpoScriptsReason = if ($userGpoScripts.Count -eq 0 -and $userGpoMissingCategories.Count -gt 0) {
    "レジストリキーが存在しないカテゴリ: $($userGpoMissingCategories -join ', ')。" +
    'ユーザー側のGPOログオン・ログオフスクリプトが設定されていないか、' +
    'ポリシーがまだこのキーへ反映されていない可能性があります。' +
    'スクリプトの設定は C:\Windows\System32\GroupPolicy\User\Scripts\ を確認してください。'
} else { $null }
$f = Save-Json "10i2_user_gpo_logon_scripts.json" $userGpoScripts $userGpoScriptsReason
Append-Index "10i2. ユーザー側 GPO ログオン・ログオフスクリプト (HKCU)" $f

# ─────────────────────────────────────────────
# 10i3. フォルダリダイレクト
# ─────────────────────────────────────────────
$shellFolderKeys = @(
    'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders',
    'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders',
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders',
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders'
)
$folderRedirect = [ordered]@{}
foreach ($key in $shellFolderKeys) {
    $props = Get-ItemProperty $key -ErrorAction SilentlyContinue
    if (-not $props) { continue }
    $label = ($key -replace 'HKCU:\\','User\' -replace 'HKLM:\\','Machine\') -replace '.*\\Explorer\\',''
    $entries = [ordered]@{}
    $props.PSObject.Properties | Where-Object { $_.Name -notmatch '^PS' } | ForEach-Object {
        # ネットワークパス（\\）または環境変数展開後がネットワークパスになるものを優先記録
        $entries[$_.Name] = $_.Value
    }
    $folderRedirect[$label] = $entries
}
# グループポリシーによるフォルダリダイレクト設定（ポリシー優先値）
$fdPolicy = Get-ItemProperty 'HKCU:\Software\Policies\Microsoft\Windows\System\Fdeploy' -ErrorAction SilentlyContinue
if ($fdPolicy) {
    $policyEntries = [ordered]@{}
    $fdPolicy.PSObject.Properties | Where-Object { $_.Name -notmatch '^PS' } | ForEach-Object {
        $policyEntries[$_.Name] = $_.Value
    }
    $folderRedirect['Policy\Fdeploy'] = $policyEntries
}
$f = Save-Json "10i3_folder_redirection.json" $folderRedirect
Append-Index "10i3. フォルダリダイレクト" $f

# ─────────────────────────────────────────────
# 10i4. デフォルトユーザープロファイルの状態
# ─────────────────────────────────────────────
$defaultProfilePath = "$env:SystemDrive\Users\Default"
$defaultProfileInfo = [ordered]@{
    Path    = $defaultProfilePath
    Exists  = (Test-Path $defaultProfilePath)
    # NTUSER.DAT の更新日（カスタマイズの目安）
    NtuserDatLastWrite = if (Test-Path "$defaultProfilePath\NTUSER.DAT") {
        (Get-Item "$defaultProfilePath\NTUSER.DAT" -Force).LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss')
    } else { $null }
    # デスクトップ・スタートアップフォルダのファイル一覧
    DesktopFiles = @(Get-ChildItem "$defaultProfilePath\Desktop" -ErrorAction SilentlyContinue |
        Select-Object Name, Extension, LastWriteTime | ForEach-Object {
            [ordered]@{ Name=$_.Name; Extension=$_.Extension; LastWrite=$_.LastWriteTime.ToString('yyyy-MM-dd') }
        })
    StartupFiles = @(Get-ChildItem "$defaultProfilePath\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup" -ErrorAction SilentlyContinue |
        Select-Object Name | ForEach-Object { $_.Name })
    # レジストリハイブ内の Run キー（管理者なら読める可能性あり）
    RunRegistry  = @()
}
# デフォルトユーザーのレジストリハイブを一時ロード
$hiveLoaded = $false
if ($isAdmin -and (Test-Path "$defaultProfilePath\NTUSER.DAT")) {
    $tmpHive = 'HKLM:\TempDefaultHive'
    $ret = reg load 'HKLM\TempDefaultHive' "$defaultProfilePath\NTUSER.DAT" 2>&1
    if ($LASTEXITCODE -eq 0) {
        $hiveLoaded = $true
        $runProps = Get-ItemProperty "$tmpHive\SOFTWARE\Microsoft\Windows\CurrentVersion\Run" -ErrorAction SilentlyContinue
        if ($runProps) {
            $runProps.PSObject.Properties | Where-Object { $_.Name -notmatch '^PS' } | ForEach-Object {
                $defaultProfileInfo.RunRegistry += [ordered]@{ Name=$_.Name; Command=$_.Value }
            }
        }
        [gc]::Collect()
        reg unload 'HKLM\TempDefaultHive' 2>&1 | Out-Null
    }
}
$f = Save-Json "10i4_default_user_profile.json" $defaultProfileInfo
Append-Index "10i4. デフォルトユーザープロファイル状態" $f

# ─────────────────────────────────────────────
# 10i5. ドメインアカウントのログオンスクリプトパス
# ─────────────────────────────────────────────
$domainLogonScript = [ordered]@{}

# 現在ログイン中のユーザーの net user 情報
$netUserLines = (net user $env:USERNAME 2>&1)
$scriptLine = $netUserLines | Where-Object { $_ -match 'スクリプト|Script' } | Select-Object -First 1
$domainLogonScript['CurrentUser_NetUser_ScriptPath'] = if ($scriptLine) {
    ($scriptLine -split '\s{2,}')[-1].Trim()
} else { $null }

# ドメイン参加の場合は net user /domain も試行
if ((Get-CimInstance Win32_ComputerSystem -ErrorAction SilentlyContinue).PartOfDomain) {
    $netUserDomainLines = (net user $env:USERNAME /domain 2>&1)
    $dScriptLine = $netUserDomainLines | Where-Object { $_ -match 'スクリプト|Script' } | Select-Object -First 1
    $domainLogonScript['CurrentUser_Domain_ScriptPath'] = if ($dScriptLine) {
        ($dScriptLine -split '\s{2,}')[-1].Trim()
    } else { $null }
}

# Netlogon スクリプトフォルダの確認（ドメインコントローラーのNetlogon共有経由）
$domainLogonScript['NetlogonShare'] = if (Test-Path '\\.\SYSVOL' -ErrorAction SilentlyContinue) { '\\.\SYSVOL' } else { $null }

$f = Save-Json "10i5_domain_logon_script.json" $domainLogonScript
Append-Index "10i5. ドメインアカウントのログオンスクリプトパス" $f

# ─────────────────────────────────────────────
# 10i6. プロファイルの上書きポリシー
# ─────────────────────────────────────────────
$profilePolicyKeys = @(
    'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon',
    'HKLM:\SOFTWARE\Policies\Microsoft\Windows\System',
    'HKCU:\Software\Policies\Microsoft\Windows\System'
)
$profilePolicy = [ordered]@{}
foreach ($key in $profilePolicyKeys) {
    $props = Get-ItemProperty $key -ErrorAction SilentlyContinue
    if (-not $props) { continue }
    $label = $key -replace 'HKLM:\\','Machine\' -replace 'HKCU:\\','User\'
    $relevant = [ordered]@{}
    $interestingNames = @(
        # ローミングプロファイル関連
        'DeleteRoamingCache','ProfileDlgTimeOut','SlowLinkTimeOut','SlowLinkUIEnabled',
        'ExcludeProfileDirs','RoamingProfileSupportEnabled',
        # ログオフ時にローカルコピーを削除
        'DeleteCachedCopies','CleanupProfiles',
        # 必須プロファイル
        'ForceUnloadHive',
        # プロファイルパス設定
        'ProfilesDirectory','DefaultUserProfile','ProfilePath',
        # グループポリシーによる制御
        'DisableForceUnload','EnableSlowLinkDetect',
        # Winlogon
        'Userinit','Shell','UserInit'
    )
    $props.PSObject.Properties | Where-Object { $_.Name -in $interestingNames } | ForEach-Object {
        $relevant[$_.Name] = $_.Value
    }
    if ($relevant.Count -gt 0) { $profilePolicy[$label] = $relevant }
}
$f = Save-Json "10i6_profile_policy.json" $profilePolicy
Append-Index "10i6. プロファイルの上書きポリシー" $f

# ─────────────────────────────────────────────
# 10j. UWF (Unified Write Filter) 状態
#      ─ キオスク・共用PC での再ログイン時リセット手段として使われる
# ─────────────────────────────────────────────
$uwfLines = @("=== UWF フィルター設定 ===")
$uwfLines += (uwfmgr.exe filter get-config 2>&1)
$uwfLines += ""
$uwfLines += "=== UWF 保護ボリューム ==="
$uwfLines += (uwfmgr.exe volume get-config all 2>&1)
$f = Save-Text "10j_uwf_status.txt" $uwfLines `
    'uwfmgr.exe がエラーを返したか出力がありませんでした。UWF（Unified Write Filter）がこのシステムで有効化されていないか、Windows の組み込み機能として追加されていない可能性があります。有効化するには「コントロールパネル → プログラムと機能 → Windows の機能の有効化または無効化」から「Unified Write Filter」を追加してください。'
Append-Index "10j. UWF (Unified Write Filter) 状態" $f

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

$monitors = Get-CimInstance Win32_VideoController -ErrorAction SilentlyContinue | ForEach-Object {
    [ordered]@{
        Name                 = $_.Name
        CurrentHorizontalRes = $_.CurrentHorizontalResolution
        CurrentVerticalRes   = $_.CurrentVerticalResolution
        CurrentRefreshRate   = $_.CurrentRefreshRate
        BitsPerPixel         = $_.CurrentBitsPerPixel
        DriverVersion        = $_.DriverVersion
        DriverDate           = if ($_.DriverDate) { $_.DriverDate.ToString('yyyy-MM-dd') } else { $null }
    }
}

# DPI スケーリング（レジストリから）
$dpiReg = Get-ItemProperty 'HKCU:\Control Panel\Desktop' -ErrorAction SilentlyContinue
$displays = [ordered]@{
    Monitors       = @($monitors)
    LogPixels      = $dpiReg.LogPixels
    Win8DpiScaling = $dpiReg.Win8DpiScaling
}
$f = Save-Json "14_display_settings.json" $displays
Append-Index "14. ディスプレイ設定" $f

# ─────────────────────────────────────────────
# 15. 共有フォルダ・ネットワークドライブ
# ─────────────────────────────────────────────
Write-Step 15 "共有・ネットワークドライブ"

# 共有フォルダ
$shares = try {
    Get-SmbShare -ErrorAction Stop | Where-Object { $_.Name -notmatch '\$$' } |
        ForEach-Object {
            [ordered]@{
                Name        = $_.Name
                Path        = $_.Path
                Description = $_.Description
                ShareState  = "$($_.ShareState)"
            }
        }
} catch {
    Write-Warning "[15] Get-SmbShare 失敗: $_"
    @()
}
$f = Save-Json "15_shared_folders.json" $shares
Append-Index "15. 共有フォルダ" $f

# ネットワークドライブ（複数手段で収集）
$netDrives = [System.Collections.Generic.List[object]]::new()

# 手段1: Get-PSDrive（現在マウント中）
try {
    Get-PSDrive -PSProvider FileSystem -ErrorAction Stop |
        Where-Object { $_.DisplayRoot -match '\\\\' } |
        ForEach-Object {
            $netDrives.Add([ordered]@{ Source='PSDrive'; DriveLetter=$_.Name; UNCPath=$_.DisplayRoot; Description=$null })
        }
} catch {}

# 手段2: WMI Win32_MappedLogicalDisk
try {
    Get-CimInstance Win32_MappedLogicalDisk -ErrorAction Stop | ForEach-Object {
        $letter = $_.Name -replace ':',''
        if (-not ($netDrives | Where-Object { $_.DriveLetter -eq $letter })) {
            $netDrives.Add([ordered]@{ Source='WMI'; DriveLetter=$letter; UNCPath=$_.ProviderName; Description=$_.Description })
        }
    }
} catch {}

# 手段3: レジストリ HKCU:\Network（永続マップ）
try {
    Get-ChildItem 'HKCU:\Network' -ErrorAction Stop | ForEach-Object {
        $props = Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue
        $letter = $_.PSChildName
        if (-not ($netDrives | Where-Object { $_.DriveLetter -eq $letter })) {
            $netDrives.Add([ordered]@{ Source='Registry'; DriveLetter=$letter; UNCPath=$props.RemotePath; Description=$props.UserName })
        }
    }
} catch {}

# 手段4: net use コマンド
$netUseRaw = (net use 2>&1) | Where-Object { $_ -match 'OK|切断|Disconnected|接続済' }
foreach ($line in $netUseRaw) {
    if ($line -match '([A-Z]):.*?(\\\\[^\s]+)') {
        $letter = $matches[1]; $unc = $matches[2]
        if (-not ($netDrives | Where-Object { $_.DriveLetter -eq $letter })) {
            $netDrives.Add([ordered]@{ Source='NetUse'; DriveLetter=$letter; UNCPath=$unc; Description=$null })
        }
    }
}

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
# 16b. デバイスマネージャー ハードウェア ID 一覧
#      Win32_PnPSignedDriver にはないHardwareID/CompatibleIDをWin32_PnPEntityから取得
# ─────────────────────────────────────────────
$allPnpDevices = Get-CimInstance Win32_PnPEntity -ErrorAction SilentlyContinue |
    Sort-Object Name | ForEach-Object {
        [ordered]@{
            Name                   = $_.Name
            DeviceID               = $_.DeviceID
            HardwareID             = @($_.HardwareID)
            CompatibleID           = @($_.CompatibleID)
            PNPClass               = $_.PNPClass
            ClassGuid              = $_.ClassGuid
            Manufacturer           = $_.Manufacturer
            Service                = $_.Service
            Status                 = $_.Status
            ConfigManagerErrorCode = $_.ConfigManagerErrorCode
        }
    }
$f = Save-Json "16b_device_hardware_ids.json" $allPnpDevices
Append-Index "16b. デバイス ハードウェア ID 一覧" $f

# ─────────────────────────────────────────────
# 16c. 専門周辺機器の詳細情報
#      キーワード/VIDフィルタで関連デバイスを絞り込み、ドライバー情報を結合
#      ペンタブレット / オーディオインターフェース / MIDIコントローラー /
#      映像キャプチャ / 配信機器 等の主要メーカーに対応
# ─────────────────────────────────────────────

# 検索キーワード（デバイス名・メーカー・サービス名・DeviceID対象）
$specializedDeviceKeywords = (
    # ペンタブレット・デジタイザー
    'wacom|tablet|pen|digitizer|stylus|bamboo|intuos|cintiq|huion|xp.pen|gaomon|ugee|screenpad|tabletpc|' +
    # オーディオインターフェース・MIDIコントローラー
    'focusrite|scarlett|clarett|universal.audio|apollo|motu|presonus|steinberg|ur[0-9]|yamaha|behringer|' +
    'roland|korg|rme|fireface|apogee|audient|ssl|m.audio|maudio|tascam|avid|mbox|midi|asio|' +
    # 映像キャプチャ・配信機器
    'elgato|blackmagic|decklink|avermedia|magewell|epiphan|capture.card|stream.deck|camlink|4k60'
)
# 専門周辺機器メーカー USB Vendor ID
# ペンタブ: Wacom=056A, Huion=28BD/2DC8, XP-Pen=0531, Gaomon=08F9
# オーディオ: Focusrite=1235, MOTU=07FD, Universal Audio=2B40, Steinberg=0499
# 映像キャプチャ: Elgato=0FD9, Blackmagic=1EDB, AVerMedia=07CA
$specializedVIDs = 'VID_056A|VID_28BD|VID_2DC8|VID_0531|VID_08F9|VID_1235|VID_07FD|VID_2B40|VID_0499|VID_0FD9|VID_1EDB|VID_07CA'

# Win32_PnPSignedDriver をメモリ上でマップ化してN+1クエリを回避
$driverLookup = @{}
Get-CimInstance Win32_PnPSignedDriver -ErrorAction SilentlyContinue |
    Where-Object { $_.DeviceID } |
    ForEach-Object { $driverLookup[$_.DeviceID] = $_ }

$specializedDeviceDetails = @(
    Get-CimInstance Win32_PnPEntity -ErrorAction SilentlyContinue |
    Where-Object {
        $searchStr = "$($_.Name) $($_.Manufacturer) $($_.Service) $($_.DeviceID) $(($_.HardwareID) -join ' ')"
        $searchStr -imatch $specializedDeviceKeywords -or $searchStr -imatch $specializedVIDs
    } |
    Sort-Object Name |
    ForEach-Object {
        $drv = $driverLookup[$_.DeviceID]
        [ordered]@{
            Name                   = $_.Name
            DeviceID               = $_.DeviceID
            HardwareID             = @($_.HardwareID)
            CompatibleID           = @($_.CompatibleID)
            PNPClass               = $_.PNPClass
            ClassGuid              = $_.ClassGuid
            Manufacturer           = $_.Manufacturer
            Service                = $_.Service
            Status                 = $_.Status
            ConfigManagerErrorCode = $_.ConfigManagerErrorCode
            Driver = if ($drv) {
                [ordered]@{
                    DriverVersion = $drv.DriverVersion
                    DriverDate    = $drv.DriverDate
                    Manufacturer  = $drv.Manufacturer
                    InfName       = $drv.InfName
                    IsSigned      = $drv.IsSigned
                }
            } else { $null }
        }
    }
)

# ベンダー固有レジストリキーを収集（型番・設定情報）
$vendorRegistry = [ordered]@{}
@(
    'HKLM:\SOFTWARE\Wacom',
    'HKLM:\SOFTWARE\Wacom\Tablet',
    'HKLM:\SOFTWARE\Tablet',
    'HKLM:\SYSTEM\CurrentControlSet\Services\WTabletServicePro',
    'HKLM:\SYSTEM\CurrentControlSet\Services\wacomvhid',
    'HKLM:\SYSTEM\CurrentControlSet\Services\WacomTouchFilter',
    'HKCU:\SOFTWARE\Wacom'
) | ForEach-Object {
    $regPath = $_
    $props = Get-ItemProperty $regPath -ErrorAction SilentlyContinue
    if ($props) {
        $vendorRegistry[$regPath] = @(
            $props.PSObject.Properties |
            Where-Object { $_.Name -notmatch '^PS' } |
            ForEach-Object { [ordered]@{ Name = $_.Name; Value = "$($_.Value)" } }
        )
    }
}

$specializedDeviceInfo = [ordered]@{
    MatchedDeviceCount = $specializedDeviceDetails.Count
    MatchedDevices     = $specializedDeviceDetails
    VendorRegistry     = $vendorRegistry
}
$f = Save-Json "16c_specialized_devices.json" $specializedDeviceInfo
Append-Index "16c. 専門周辺機器詳細（ペンタブ／オーディオI/F／映像キャプチャ等）" $f

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
$printers = try {
    Get-Printer -ErrorAction Stop | ForEach-Object {
        [ordered]@{
            Name          = $_.Name
            DriverName    = $_.DriverName
            PortName      = $_.PortName
            Shared        = $_.Shared
            Default       = if ($_.PSObject.Properties['Default']) { $_.Default } else { $null }
            PrinterStatus = "$($_.PrinterStatus)"
        }
    }
} catch {
    Write-Warning "[17b] Get-Printer 失敗: $_"
    @()
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
# gpresult は Unicode 出力を行うため、OutputEncoding を Unicode に切り替えて文字化けを防止
$gpresultLines = @("=== gpresult /r ===")
$savedGpEnc = [Console]::OutputEncoding
[Console]::OutputEncoding = [System.Text.Encoding]::Unicode
$gpresultLines += (gpresult /r 2>&1)
[Console]::OutputEncoding = $savedGpEnc
$f = Save-Text "19_gpresult.txt" $gpresultLines
Append-Index "19. GPO 適用結果（テキスト）" $f

# gpresult HTML形式（詳細レポート）
# /user 指定なしでは実行ユーザー分のみ取得できるため、管理者時はプロファイルを列挙して全ユーザー分取得する
if ($isAdmin) {
    $gpUserProfiles = Get-WmiObject Win32_UserProfile -ErrorAction SilentlyContinue |
        Where-Object { -not $_.Special -and (Test-Path $_.LocalPath) } |
        Sort-Object LocalPath
    foreach ($gpp in $gpUserProfiles) {
        try {
            $gpNtAccount = (New-Object System.Security.Principal.SecurityIdentifier($gpp.SID)).Translate(
                [System.Security.Principal.NTAccount]).Value
            $gpSafeName  = $gpNtAccount -replace '[\\\/\s:*?"<>|]', '_'
            $gpHtmlPath  = Join-Path $outDir "19b_gpresult_$gpSafeName.html"
            gpresult /user $gpNtAccount /h $gpHtmlPath /f 2>&1 | Out-Null
            if (Test-Path $gpHtmlPath) {
                Append-Index "19b. GPO 適用結果 HTML ($gpNtAccount)" $gpHtmlPath
            }
        } catch {
            # SID を NTAccount に変換できない（削除済みアカウント等）はスキップ
        }
    }
} else {
    $gpresultHtml = Join-Path $outDir "19b_gpresult.html"
    gpresult /h $gpresultHtml /f 2>&1 | Out-Null
    if (Test-Path $gpresultHtml) {
        Append-Index "19b. GPO 適用結果（HTML詳細）" $gpresultHtml
    }
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

$allPolicies = [ordered]@{
    ComputerPolicies       = @(Get-RegistryPolicies 'HKLM:\SOFTWARE\Policies')
    ComputerPoliciesLegacy = @(Get-RegistryPolicies 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies')
    UserPoliciesPerUser    = [ordered]@{}
}

# 全ユーザーの HKCU ポリシーを収集する
# ログイン中ユーザーは HKU に hive が既にロードされているため直接スキャンし、
# ログアウト済みユーザーは NTUSER.DAT を reg load でマウントしてスキャン後に unload する
# （reg load には管理者権限が必要）
if ($isAdmin) {
    $regUserProfiles = Get-WmiObject Win32_UserProfile -ErrorAction SilentlyContinue |
        Where-Object { -not $_.Special -and (Test-Path $_.LocalPath) }
    foreach ($rup in $regUserProfiles) {
        $rupSid     = $rup.SID
        $rupHkuBase = "Registry::HKEY_USERS\$rupSid"
        $wasLoaded  = Test-Path $rupHkuBase

        if (-not $wasLoaded) {
            $ntdatPath = Join-Path $rup.LocalPath 'NTUSER.DAT'
            if (Test-Path $ntdatPath) {
                reg load "HKU\$rupSid" $ntdatPath 2>&1 | Out-Null
            }
        }

        try {
            $rupAccount = (New-Object System.Security.Principal.SecurityIdentifier($rupSid)).Translate(
                [System.Security.Principal.NTAccount]).Value
        } catch {
            $rupAccount = $rupSid
        }

        $rupPolicies1 = @(Get-RegistryPolicies "$rupHkuBase\SOFTWARE\Policies")
        $rupPolicies2 = @(Get-RegistryPolicies "$rupHkuBase\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies")
        $allPolicies.UserPoliciesPerUser[$rupAccount] = [ordered]@{
            UserPolicies       = $rupPolicies1
            UserPoliciesLegacy = $rupPolicies2
        }

        if (-not $wasLoaded) {
            # アンロード前に GC でレジストリハンドルを解放しておく
            [GC]::Collect()
            [GC]::WaitForPendingFinalizers()
            reg unload "HKU\$rupSid" 2>&1 | Out-Null
        }
    }
} else {
    # 非管理者の場合は実行ユーザー分のみ
    $allPolicies.UserPoliciesPerUser['(current user)'] = [ordered]@{
        UserPolicies       = @(Get-RegistryPolicies 'HKCU:\SOFTWARE\Policies')
        UserPoliciesLegacy = @(Get-RegistryPolicies 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies')
    }
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
# 20. スクリプト・バッチファイルの実体収集
# ─────────────────────────────────────────────
Write-Step 20 "スクリプト・バッチファイルの実体収集"

$scriptCopyDir = Join-Path $outDir "20_scripts"
New-Item -ItemType Directory -Path $scriptCopyDir -Force | Out-Null

$scriptExtensions = @('.bat', '.cmd', '.ps1', '.vbs', '.wsf', '.js', '.adm', '.admx')
$scriptManifest   = [System.Collections.Generic.List[object]]::new()

function Copy-ScriptFile {
    param([string]$sourcePath, [string]$category, [string]$reason)
    if (-not $sourcePath) { return }
    $sourcePath = [System.Environment]::ExpandEnvironmentVariables($sourcePath)
    if (-not (Test-Path $sourcePath -PathType Leaf -ErrorAction SilentlyContinue)) { return }

    $fileName    = Split-Path $sourcePath -Leaf
    $safeCategory = $category -replace '[\\/:*?"<>|]', '_'
    $destName    = "${safeCategory}_${fileName}"
    $destPath    = Join-Path $scriptCopyDir $destName

    $i = 1
    while (Test-Path $destPath) {
        $ext  = [System.IO.Path]::GetExtension($fileName)
        $base = [System.IO.Path]::GetFileNameWithoutExtension($fileName)
        $destName = "${safeCategory}_${base}_${i}${ext}"
        $destPath = Join-Path $scriptCopyDir $destName
        $i++
    }

    try {
        Copy-Item $sourcePath $destPath -ErrorAction Stop
        $script:scriptManifest.Add([ordered]@{
            Category   = $category
            Reason     = $reason
            SourcePath = $sourcePath
            CopiedAs   = $destName
            SizeBytes  = (Get-Item $destPath).Length
            LastWrite  = (Get-Item $sourcePath).LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss')
        })
    } catch {
        $script:scriptManifest.Add([ordered]@{
            Category   = $category
            Reason     = $reason
            SourcePath = $sourcePath
            CopiedAs   = $null
            Error      = $_.Exception.Message
        })
    }
}

# ── (a) GPO スクリプトディレクトリをスキャン ──
$gpoDirs = @(
    @{ Path = "$env:SystemRoot\System32\GroupPolicy\Machine\Scripts\Startup";  Category = 'GPO_Machine_Startup'  }
    @{ Path = "$env:SystemRoot\System32\GroupPolicy\Machine\Scripts\Shutdown"; Category = 'GPO_Machine_Shutdown' }
    @{ Path = "$env:SystemRoot\System32\GroupPolicy\User\Scripts\Logon";       Category = 'GPO_User_Logon'       }
    @{ Path = "$env:SystemRoot\System32\GroupPolicy\User\Scripts\Logoff";      Category = 'GPO_User_Logoff'      }
    @{ Path = "$env:SystemRoot\SysWOW64\GroupPolicy\Machine\Scripts\Startup";  Category = 'GPO_Machine_Startup'  }
    @{ Path = "$env:SystemRoot\SysWOW64\GroupPolicy\Machine\Scripts\Shutdown"; Category = 'GPO_Machine_Shutdown' }
)
foreach ($gpoDir in $gpoDirs) {
    if (-not (Test-Path $gpoDir.Path -ErrorAction SilentlyContinue)) { continue }
    Get-ChildItem $gpoDir.Path -Recurse -File -ErrorAction SilentlyContinue |
        Where-Object { $scriptExtensions -contains $_.Extension.ToLower() } |
        ForEach-Object { Copy-ScriptFile $_.FullName $gpoDir.Category 'GPO スクリプトディレクトリ' }
}

# ── (b) スタートアップフォルダをスキャン ──
$startupScanFolders = @(
    @{ Path = [System.Environment]::GetFolderPath('CommonStartup'); Category = 'Startup_AllUsers'    }
    @{ Path = [System.Environment]::GetFolderPath('Startup');       Category = 'Startup_CurrentUser' }
)
foreach ($sf in $startupScanFolders) {
    if (-not (Test-Path $sf.Path -ErrorAction SilentlyContinue)) { continue }
    Get-ChildItem $sf.Path -Recurse -File -ErrorAction SilentlyContinue |
        Where-Object { $scriptExtensions -contains $_.Extension.ToLower() } |
        ForEach-Object { Copy-ScriptFile $_.FullName $sf.Category 'スタートアップフォルダ' }
}

# ── (c) GPO スクリプトレジストリから参照先ファイルをコピー ──
foreach ($category in @('Logon','Logoff','Startup','Shutdown')) {
    $gpoRegPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\Scripts\$category"
    if (-not (Test-Path $gpoRegPath -ErrorAction SilentlyContinue)) { continue }
    Get-ChildItem $gpoRegPath -ErrorAction SilentlyContinue | ForEach-Object {
        Get-ChildItem $_.PSPath -ErrorAction SilentlyContinue | ForEach-Object {
            $entry = Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue
            if ($entry.Script) {
                Copy-ScriptFile $entry.Script "GPO_Reg_$category" 'GPO スクリプトレジストリ参照'
            }
        }
    }
}

# ── (d) スケジュールタスクのアクションがスクリプトを参照している場合 ──
Get-ScheduledTask -ErrorAction SilentlyContinue |
    Where-Object { $_.TaskPath -notlike '\Microsoft\*' } |
    ForEach-Object {
        $taskName = $_.TaskName
        $_.Actions | ForEach-Object {
            $exe = $_.Execute
            if ($exe -and ($scriptExtensions -contains [System.IO.Path]::GetExtension($exe).ToLower())) {
                Copy-ScriptFile $exe 'ScheduledTask' "タスクスケジューラ: $taskName"
            }
            # 引数にスクリプトパスが含まれる場合（例: powershell.exe -File C:\foo.ps1）
            if ($_.Arguments -match '"?([A-Za-z]:\\[^"]+\.(ps1|bat|cmd|vbs|wsf))"?') {
                Copy-ScriptFile $matches[1] 'ScheduledTask' "タスクスケジューラ引数: $taskName"
            }
        }
    }

$f = Save-Json "20_scripts_manifest.json" $scriptManifest
Append-Index "20. スクリプト・バッチファイル実体（コピー先: 20_scripts/）" $f

# ─────────────────────────────────────────────
# 21. システム情報（msinfo32 システムの要約）
# ─────────────────────────────────────────────
Write-Step 21 "システム情報（msinfo32 システムの要約）"

try {
    $tempReport = Join-Path $env:TEMP "msinfo32_report_$timestamp.txt"
    $proc = Start-Process -FilePath "$env:SystemRoot\System32\msinfo32.exe" `
        -ArgumentList "/report `"$tempReport`"" -PassThru -WindowStyle Hidden -ErrorAction Stop
    $waited = $proc.WaitForExit(60000)
    if (-not $waited) { try { $proc.Kill() } catch {} }

    if (Test-Path $tempReport) {
        $rawLines = [System.IO.File]::ReadAllLines($tempReport, [System.Text.Encoding]::Unicode)

        # システムの要約セクション（日本語・英語両対応）を抽出
        $inSection = $false
        $sectionData = [ordered]@{}
        foreach ($line in $rawLines) {
            if ($line -match '^\[(システムの要約|System Summary)\]') {
                $inSection = $true
                continue
            }
            if ($inSection -and $line -match '^\[') { break }
            if ($inSection -and $line -match '^(.+?)\t(.*)$') {
                $key = $matches[1].Trim()
                $val = $matches[2].Trim()
                # 列ヘッダー行（"項目" / "Item"）はスキップ
                if ($key -ne '項目' -and $key -ne 'Item') {
                    $sectionData[$key] = $val
                }
            }
        }

        Remove-Item $tempReport -Force -ErrorAction SilentlyContinue

        if ($sectionData.Count -gt 0) {
            $f = Save-Json "21_msinfo32_summary.json" $sectionData
        } else {
            $f = Save-Json "21_msinfo32_summary.json" $null `
                'msinfo32 の出力からシステムの要約セクションを取得できませんでした。'
        }
    } else {
        $f = Save-Json "21_msinfo32_summary.json" $null `
            "msinfo32 /report の出力ファイルが生成されませんでした。$env:SystemRoot\System32\msinfo32.exe が存在するか確認してください。"
    }
} catch {
    $errMsg = "$_"
    $f = Save-Json "21_msinfo32_summary.json" $null "msinfo32 の実行に失敗しました: $errMsg"
}
Append-Index "21. システム情報（msinfo32 システムの要約）" $f

# ─────────────────────────────────────────────
# 22. 完了
# ─────────────────────────────────────────────
Write-Step 22 "完了"

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
# HTML レポート生成（-GenerateHtml 指定時）
# ─────────────────────────────────────────────
$htmlPath = $null
if ($GenerateHtml) {
    Write-Host " HTML レポート生成中..." -ForegroundColor Cyan

    $navSb  = [System.Text.StringBuilder]::new()
    $secsSb = [System.Text.StringBuilder]::new()
    [void]$navSb.Append('<ul>')

    $catName = ''
    foreach ($idxLine in $indexLines) {
        if ($idxLine -match '^## (.+)') { $catName = $matches[1] }
        if ($idxLine -match '^\s+ファイル:\s+(.+)' -and $catName -ne '') {
            $fn  = $matches[1].Trim()
            $sid = $fn -replace '[^a-zA-Z0-9]', '_'
            $encCat = [System.Net.WebUtility]::HtmlEncode($catName)
            $encFn  = [System.Net.WebUtility]::HtmlEncode($fn)
            [void]$navSb.Append("<li><a href='#$sid' title='$encCat'>$encCat</a></li>")

            $fp  = Join-Path $outDir $fn
            $ext = [System.IO.Path]::GetExtension($fn).ToLower()
            [void]$secsSb.Append("<section id='$sid'>")
            [void]$secsSb.Append("<h2>$encCat</h2>")
            [void]$secsSb.Append("<p class='filename'>$encFn</p>")
            if (-not (Test-Path $fp -PathType Leaf)) {
                [void]$secsSb.Append("<p class='empty'>ファイルが見つかりません</p>")
            } elseif ($ext -eq '.html') {
                [void]$secsSb.Append("<p><a href='$encFn' target='_blank'>別ウィンドウで開く ↗</a></p>")
                [void]$secsSb.Append("<iframe src='$encFn' style='width:100%;height:600px;border:1px solid #444;'></iframe>")
            } else {
                try   { $raw = [System.IO.File]::ReadAllText($fp, [System.Text.Encoding]::UTF8) }
                catch { $raw = '' }
                $cls = if ($ext -eq '.json') { 'json' } else { 'text' }
                [void]$secsSb.Append("<pre class='$cls'>$([System.Net.WebUtility]::HtmlEncode($raw))</pre>")
            }
            [void]$secsSb.Append('</section>')
        }
    }
    [void]$navSb.Append('</ul>')

    $navHtml  = $navSb.ToString()
    $bodyHtml = $secsSb.ToString()

    $metaHost  = [System.Net.WebUtility]::HtmlEncode($env:COMPUTERNAME)
    $metaUser  = [System.Net.WebUtility]::HtmlEncode($env:USERNAME)
    $metaAdmin = if ($isAdmin) { '管理者権限あり' } else { '管理者権限なし' }
    $metaDate  = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'

    $htmlReport = @"
<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Windows 環境調査レポート - $metaHost</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Segoe UI',sans-serif;background:#1e1e1e;color:#d4d4d4;display:flex;min-height:100vh}
#sidebar{width:260px;min-width:260px;background:#252526;position:sticky;top:0;height:100vh;overflow-y:auto}
#sidebar h1{font-size:13px;font-weight:600;color:#9cdcfe;padding:12px 16px 8px;border-bottom:1px solid #3e3e42}
#sidebar .meta{font-size:11px;color:#858585;padding:8px 16px 10px;border-bottom:1px solid #3e3e42;line-height:1.7}
#sidebar ul{list-style:none;padding-bottom:16px}
#sidebar ul li a{display:block;padding:3px 16px;font-size:12px;color:#cccccc;text-decoration:none;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
#sidebar ul li a:hover{background:#2a2d2e;color:#fff}
#main{flex:1;padding:24px;overflow-x:hidden;min-width:0}
section{margin-bottom:32px;border:1px solid #3e3e42;border-radius:4px;overflow:hidden}
section h2{font-size:13px;font-weight:600;background:#2d2d30;color:#9cdcfe;padding:9px 16px;border-bottom:1px solid #3e3e42}
.filename{font-size:11px;color:#6a9955;padding:3px 16px;background:#252526;border-bottom:1px solid #3e3e42}
pre{padding:14px 16px;font-size:12px;font-family:'Cascadia Code','Consolas',monospace;white-space:pre-wrap;word-break:break-all;line-height:1.5;max-height:600px;overflow-y:auto;background:#1e1e1e}
pre.json{color:#9cdcfe}
pre.text{color:#d4d4d4}
.empty{padding:10px 16px;color:#858585;font-style:italic}
a{color:#4ec9b0}
iframe{display:block}
</style>
</head>
<body>
<nav id="sidebar">
<h1>Windows 環境調査</h1>
<div class="meta">ホスト: $metaHost<br>ユーザー: $metaUser<br>権限: $metaAdmin<br>収集日時: $metaDate</div>
$navHtml
</nav>
<main id="main">
$bodyHtml
</main>
</body>
</html>
"@

    $htmlPath = Join-Path $outDir "00_REPORT.html"
    [System.IO.File]::WriteAllText($htmlPath, $htmlReport, [System.Text.Encoding]::UTF8)
}

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
if ($htmlPath) {
    Write-Host " HTML レポート: $htmlPath"
}
Write-Host "=" * 60 -ForegroundColor Green
