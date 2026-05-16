#Requires -Version 5.1
<#
.SYNOPSIS
    WindowsEnvAuditフォルダをMarkdownパラメータシートに変換する

.PARAMETER AuditPath
    WindowsEnvAuditフォルダのパス。省略時はスクリプトと同じ場所の
    WindowsEnvAudit* フォルダを全て処理する。

.PARAMETER OutputPath
    出力先ディレクトリ。省略時はAuditPathの親フォルダ。

.EXAMPLE
    .\Create-ParameterSheet.ps1
    .\Create-ParameterSheet.ps1 -AuditPath "C:\Audit\WindowsEnvAudit_20260515"
    .\Create-ParameterSheet.ps1 -AuditPath ".\WindowsEnvAudit" -OutputPath "D:\Reports"
#>
[CmdletBinding()]
param(
    [string]$AuditPath  = '',
    [string]$OutputPath = ''
)

Set-StrictMode -Off
$ErrorActionPreference = 'SilentlyContinue'

# ─────────────────────────────────────────────────
# ヘルパー関数
# ─────────────────────────────────────────────────

function Read-JsonSafe {
    param([string]$Path)
    if (-not (Test-Path $Path)) { return $null }
    try {
        $raw = Get-Content -Path $Path -Encoding UTF8 -Raw
        $obj = $raw | ConvertFrom-Json
        if ($obj -is [psobject] -and $obj.PSObject.Properties['_status']) { return $null }
        return $obj
    } catch { return $null }
}

function ConvertFrom-JsonDate {
    param([object]$Value)
    if ($null -eq $Value) { return '-' }
    $s = "$Value"
    if ($s -match '/Date\((\d+)\)/') {
        try {
            $dt = [DateTimeOffset]::FromUnixTimeMilliseconds([long]$Matches[1])
            return $dt.LocalDateTime.ToString('yyyy-MM-dd HH:mm')
        } catch { return $s }
    }
    try {
        return ([datetime]::Parse($s)).ToString('yyyy-MM-dd HH:mm')
    } catch { return $s }
}

function fv {
    param([object]$Value, [string]$Default = '-')
    if ($null -eq $Value) { return $Default }
    $s = "$Value".Trim()
    if ($s -eq '') { return $Default }
    return $s
}

function fesc {
    param([object]$Value, [string]$Default = '-', [int]$MaxLen = 0)
    if ($null -eq $Value) { return $Default }
    $s = "$Value".Trim()
    if ($s -eq '') { return $Default }
    $s = $s -replace '\r?\n', ' '
    $s = $s -replace '\|', '\|'
    if ($MaxLen -gt 0 -and $s.Length -gt $MaxLen) { $s = $s.Substring(0, $MaxLen) + '...' }
    return $s
}

function Bool-Str {
    param([object]$Value, [string]$TrueLabel = '有効', [string]$FalseLabel = '無効')
    if ($null -eq $Value) { return '-' }
    $s = "$Value"
    if ($s -eq 'True' -or $s -eq '1') { return $TrueLabel }
    if ($s -eq 'False' -or $s -eq '0') { return $FalseLabel }
    return $s
}

# ─────────────────────────────────────────────────
# 処理対象フォルダの特定
# ─────────────────────────────────────────────────

$foldersToProcess = @()

if (-not [string]::IsNullOrEmpty($AuditPath)) {
    if (-not (Test-Path $AuditPath)) {
        Write-Error "指定されたフォルダが存在しません: $AuditPath"
        exit 1
    }
    $foldersToProcess = @($AuditPath)
} else {
    $scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
    $foldersToProcess = @(
        Get-ChildItem -Path $scriptDir -Directory -Filter 'WindowsEnvAudit*' |
        Sort-Object Name |
        ForEach-Object { $_.FullName }
    )
    if ($foldersToProcess.Count -eq 0) {
        Write-Error "WindowsEnvAudit* フォルダが見つかりません。-AuditPath で指定してください。"
        exit 1
    }
}

Write-Host "$($foldersToProcess.Count) 件のフォルダを処理します"

# ─────────────────────────────────────────────────
# 各フォルダを処理
# ─────────────────────────────────────────────────

foreach ($currentAuditPath in $foldersToProcess) {

    Write-Host ""
    Write-Host "処理中: $currentAuditPath"

    # --- メタデータ読み込み (00_INDEX.md) ---

    $collectedAt = '-'
    $collectedBy = '-'
    $hasAdmin    = '-'

    $indexPath = Join-Path $currentAuditPath '00_INDEX.md'
    if (Test-Path $indexPath) {
        $indexLines = Get-Content -Path $indexPath -Encoding UTF8
        foreach ($line in $indexLines) {
            if ($line -match '^生成日時[:：]\s*(.+)')   { $collectedAt = $Matches[1].Trim() }
            if ($line -match '^実行ユーザー[:：]\s*(.+)') { $collectedBy = $Matches[1].Trim() }
            if ($line -match '^管理者権限[:：]\s*(.+)')   {
                $hasAdmin = if ($Matches[1].Trim() -eq 'True') { 'あり' } else { 'なし' }
            }
        }
        if ($collectedAt -match '^(\d{1,2})/(\d{1,2})/(\d{4})\s+(.+)') {
            $collectedAt = '{0}-{1:D2}-{2:D2} {3}' -f $Matches[3], [int]$Matches[1], [int]$Matches[2], $Matches[4]
        }
    } else {
        $folderName = Split-Path $currentAuditPath -Leaf
        if ($folderName -match '_(\d{4})(\d{2})(\d{2})_(\d{2})(\d{2})(\d{2})$') {
            $collectedAt = '{0}-{1}-{2} {3}:{4}:{5}' -f $Matches[1],$Matches[2],$Matches[3],$Matches[4],$Matches[5],$Matches[6]
        }
    }

    # --- PC名取得と出力ファイル決定 ---

    $sysInfo      = Read-JsonSafe (Join-Path $currentAuditPath '01_system_info.json')
    $computerName = if ($sysInfo) { fv $sysInfo.Computer.Name 'UNKNOWN' } else { 'UNKNOWN' }

    $dateStr  = (Get-Date).ToString('yyyyMMdd')
    $destDir  = if (-not [string]::IsNullOrEmpty($OutputPath)) { $OutputPath } else { Split-Path $currentAuditPath -Parent }
    $outFile  = Join-Path $destDir "${computerName}_ParameterSheet_${dateStr}.md"

    # --- Markdown ビルド用リスト ---

    $lines = [System.Collections.Generic.List[string]]::new()

    function ln     { param([string]$s = '') $script:lines.Add($s) }
    function h2     { param([string]$t) $script:lines.Add(''); $script:lines.Add("## $t") }
    function h3     { param([string]$t) $script:lines.Add(''); $script:lines.Add("### $t") }
    function nodata { $script:lines.Add(''); $script:lines.Add('> データなし（ファイルが存在しないか、収集時に失敗）') }

    # ═══════════════════════════════════════════════════════
    # ヘッダー
    # ═══════════════════════════════════════════════════════

    ln '# Windows 環境パラメータシート'
    ln ''
    ln '| 項目 | 値 |'
    ln '|------|-----|'
    ln "| PC名 | $(fesc $computerName) |"
    ln "| 収集日時 | $(fesc $collectedAt) |"
    ln "| 収集者 | $(fesc $collectedBy) |"
    ln "| 管理者権限 | $hasAdmin |"
    if ($sysInfo) {
        $dom = fv $sysInfo.Computer.Domain
        $wg  = fv $sysInfo.Computer.Workgroup
        $domStr = if ($dom -eq 'WORKGROUP') { "ワークグループ: $wg" } else { "ドメイン: $dom" }
        ln "| ドメイン/WG | $(fesc $domStr) |"
    }
    ln ''
    ln '---'

    # ═══════════════════════════════════════════════════════
    # 1. システム情報
    # ═══════════════════════════════════════════════════════

    h2 '1. システム情報'

    if ($sysInfo) {
        $os   = $sysInfo.OS
        $cpu  = if ($sysInfo.CPU -and @($sysInfo.CPU).Count -gt 0) { @($sysInfo.CPU)[0] } else { $null }
        $bios = $sysInfo.BIOS
        $mb   = $sysInfo.Motherboard
        $pc   = $sysInfo.Computer

        h3 'OS'
        ln '| 項目 | 値 |'
        ln '|------|-----|'
        ln "| OS名 | $(fesc $os.Caption) |"
        ln "| バージョン | $(fesc $os.Version) (Build $(fesc $os.BuildNumber)) |"
        ln "| アーキテクチャ | $(fesc $os.OSArchitecture) |"

        h3 'ハードウェア'
        ln '| 項目 | 値 |'
        ln '|------|-----|'
        ln "| メーカー/モデル | $(fesc $pc.Manufacturer) / $(fesc $pc.Model) |"
        if ($cpu) {
            $cores   = fv $cpu.NumberOfCores
            $threads = fv $cpu.NumberOfLogicalProcessors
            $mhz     = fv $cpu.MaxClockSpeedMHz
            ln "| CPU | $(fesc $cpu.Name) (${cores}コア/${threads}スレッド) |"
            ln "| クロック | ${mhz} MHz |"
        }
        ln "| RAM | $(fv $pc.TotalPhysicalMemoryGB) GB |"

        h3 'GPU'
        ln '| # | GPU名 | VRAM | ドライバー |'
        ln '|---|-------|------|-----------|'
        $gpuList = @($sysInfo.GPU)
        if ($gpuList.Count -gt 0) {
            $i = 1
            foreach ($g in $gpuList) {
                ln "| $i | $(fesc $g.Name) | $(fv $g.AdapterRAMGB) GB | $(fesc $g.DriverVersion) |"
                $i++
            }
        } else { ln '| - | (GPUなし) | - | - |' }

        $specDevs = Read-JsonSafe (Join-Path $currentAuditPath '16c_specialized_devices.json')
        h3 '専用デバイス'
        if ($specDevs -and $specDevs.MatchedDevices -and @($specDevs.MatchedDevices).Count -gt 0) {
            ln '| デバイス名 | クラス | メーカー | ドライバーバージョン | INF |'
            ln '|-----------|--------|---------|------------------|-----|'
            foreach ($dev in @($specDevs.MatchedDevices)) {
                $dv  = if ($dev.Driver) { fesc $dev.Driver.DriverVersion } else { '-' }
                $inf = if ($dev.Driver) { fesc $dev.Driver.InfName } else { '-' }
                ln "| $(fesc $dev.Name) | $(fesc $dev.PNPClass) | $(fesc $dev.Manufacturer) | $dv | $inf |"
            }
        } else { ln '(専用デバイスなし)' }

        h3 'BIOS / セキュアブート'
        ln '| 項目 | 値 |'
        ln '|------|-----|'
        ln "| BIOSメーカー | $(fesc $bios.Manufacturer) |"
        ln "| BIOSバージョン | $(fesc $bios.Version) (SMBIOS $(fesc $bios.SMBIOSVersion)) |"
        ln "| リリース日 | $(ConvertFrom-JsonDate $bios.ReleaseDate) |"
        $secBoot = if ($null -eq $bios.SecureBoot) { '不明' } elseif ($bios.SecureBoot) { '有効' } else { '無効' }
        ln "| セキュアブート | $secBoot |"
        $mbProduct = "$(fv $mb.Manufacturer) $(fv $mb.Product)".Trim()
        ln "| マザーボード製品 | $(fesc $mbProduct) |"
    } else { nodata }

    # ═══════════════════════════════════════════════════════
    # 2. ストレージ
    # ═══════════════════════════════════════════════════════

    h2 '2. ストレージ'

    $disks = Read-JsonSafe (Join-Path $currentAuditPath '02_disk_partitions.json')
    h3 '物理ディスク'
    if ($disks) {
        ln '| # | モデル | 容量 (GB) | 接続 | タイプ |'
        ln '|---|--------|-----------|------|--------|'
        foreach ($d in @($disks)) {
            ln "| $(fv $d.DiskNumber) | $(fesc $d.FriendlyName) | $(fv $d.SizeGB) | $(fesc $d.BusType) | $(fesc $d.PartitionStyle) |"
        }
    } else { nodata }

    $logDisks  = Read-JsonSafe (Join-Path $currentAuditPath '02b_logical_disks.json')
    $bitlocker = Read-JsonSafe (Join-Path $currentAuditPath '10e_bitlocker.json')

    $blTable = @{}
    if ($bitlocker) {
        foreach ($b in @($bitlocker)) { $blTable[$b.MountPoint] = $b }
    }

    h3 '論理ドライブ'
    if ($logDisks) {
        $ldFiltered = @($logDisks) | Where-Object { $_.DriveType -ne 'CDRom' } | Sort-Object DeviceID
        ln '| ドライブ | FS | ラベル | サイズ (GB) | 空き (GB) | 使用率 | BitLocker |'
        ln '|----------|-----|--------|------------|-----------|--------|-----------|'
        foreach ($ld in $ldFiltered) {
            $bl      = $blTable[$ld.DeviceID]
            $blStr   = if ($bl) { "$(fesc $bl.ProtectionStatus) ($(fesc $bl.EncryptionMethod))" } else { '-' }
            $label   = fesc $ld.VolumeName '(なし)'
            $usedPct = if ($null -ne $ld.UsedPercent) { "$($ld.UsedPercent) %" } else { '-' }
            ln "| $(fesc $ld.DeviceID) | $(fesc $ld.FileSystem) | $label | $(fv $ld.SizeGB) | $(fv $ld.FreeGB) | $usedPct | $blStr |"
        }
    } else { nodata }

    $vmem = Read-JsonSafe (Join-Path $currentAuditPath '12b_virtual_memory.json')
    h3 '仮想メモリ (ページファイル)'
    if ($vmem) {
        ln "自動管理: $(Bool-Str $vmem.AutomaticManagedPagefile)"
        ln ''
        $pfList = @($vmem.PageFiles)
        if ($pfList.Count -gt 0) {
            ln '| ファイル | 割当 (MB) | 現在 (MB) | ピーク (MB) |'
            ln '|---------|-----------|-----------|------------|'
            foreach ($pf in $pfList) {
                ln "| $(fesc $pf.Name) | $(fv $pf.AllocatedBaseSizeMB) | $(fv $pf.CurrentUsageMB) | $(fv $pf.PeakUsageMB) |"
            }
        } else { ln '(ページファイルなし)' }
    } else { nodata }

    # ═══════════════════════════════════════════════════════
    # 3. ネットワーク
    # ═══════════════════════════════════════════════════════

    h2 '3. ネットワーク'

    $netRaw = Read-JsonSafe (Join-Path $currentAuditPath '09_network_adapters.json')
    if ($netRaw) {
        $adapters = @($netRaw) | Sort-Object Name
        ln '| アダプター名 | MAC | IPアドレス | CIDR | ゲートウェイ | DNS |'
        ln '|-------------|-----|-----------|------|------------|-----|'
        foreach ($a in $adapters) {
            $ipStr  = if ($a.IPv4Address)      { (@($a.IPv4Address) | Where-Object { $_ }) -join ', ' } else { '-' }
            $cidr   = if ($a.IPv4PrefixLength) { '/' + (@($a.IPv4PrefixLength)[0]) } else { '-' }
            $gwStr  = if ($a.DefaultGateway)   { (@($a.DefaultGateway) | Where-Object { $_ }) -join ', ' } else { '-' }
            $dnsStr = if ($a.DNSServers)       { (@($a.DNSServers) | Where-Object { $_ }) -join ', ' } else { '-' }
            ln "| $(fesc $a.Name) | $(fesc $a.MacAddress) | $ipStr | $cidr | $(fesc $gwStr) | $(fesc $dnsStr) |"
        }
    } else { nodata }

    # ═══════════════════════════════════════════════════════
    # 4. セキュリティ
    # ═══════════════════════════════════════════════════════

    h2 '4. セキュリティ'

    $uac = Read-JsonSafe (Join-Path $currentAuditPath '10d_uac_settings.json')
    h3 'UAC設定'
    if ($uac) {
        $adminPromptMap = @{
            0='昇格: 確認なし'; 1='昇格: 資格情報入力 (セキュアデスクトップ)';
            2='昇格: 確認 (セキュアデスクトップ)'; 3='昇格: 資格情報入力'; 4='昇格: 確認'; 5='昇格: 確認 (既定)'
        }
        $userPromptMap = @{
            0='自動拒否'; 1='資格情報入力 (セキュアデスクトップ)';
            2='確認 (セキュアデスクトップ)'; 3='資格情報入力'; 4='確認'
        }
        $adminVal = [int]$uac.ConsentPromptBehaviorAdmin
        $userVal  = [int]$uac.ConsentPromptBehaviorUser
        $adminStr = if ($adminPromptMap.ContainsKey($adminVal)) { "$($adminPromptMap[$adminVal]) ($adminVal)" } else { "$adminVal" }
        $userStr  = if ($userPromptMap.ContainsKey($userVal))   { "$($userPromptMap[$userVal]) ($userVal)"   } else { "$userVal" }
        ln '| 項目 | 値 |'
        ln '|------|-----|'
        ln "| UAC有効 | $(Bool-Str $uac.EnableLUA) |"
        ln "| 管理者の昇格動作 | $adminStr |"
        ln "| 標準ユーザーの昇格動作 | $userStr |"
        ln "| セキュアデスクトップ | $(Bool-Str $uac.PromptOnSecureDesktop) |"
    } else { nodata }

    $secProd = Read-JsonSafe (Join-Path $currentAuditPath '09g_security_products.json')
    h3 'セキュリティ製品 (SecurityCenter)'
    if ($secProd) {
        ln '| 種別 | 製品名 | 状態 |'
        ln '|------|--------|------|'
        $cats = @(
            [pscustomobject]@{Key='AntiVirus';   Label='ウイルス対策'}
            [pscustomobject]@{Key='Firewall';    Label='ファイアウォール'}
            [pscustomobject]@{Key='AntiSpyware'; Label='スパイウェア対策'}
        )
        $hasAny = $false
        foreach ($cat in $cats) {
            $items = $secProd.($cat.Key)
            if ($items -and @($items).Count -gt 0) {
                foreach ($item in @($items)) {
                    ln "| $($cat.Label) | $(fesc $item.Name) | $(fesc $item.ProductState) |"
                    $hasAny = $true
                }
            }
        }
        if (-not $hasAny) { ln '| - | (登録製品なし) | - |' }
    } else { nodata }

    $fwProfiles = Read-JsonSafe (Join-Path $currentAuditPath '09f_firewall_profiles.json')
    h3 'ファイアウォールプロファイル'
    if ($fwProfiles) {
        ln '| プロファイル | 有効 | 受信既定 | 送信既定 |'
        ln '|------------|------|---------|---------|'
        foreach ($fp in @($fwProfiles)) {
            ln "| $(fesc $fp.Name) | $(fesc $fp.Enabled) | $(fesc $fp.DefaultInboundAction) | $(fesc $fp.DefaultOutboundAction) |"
        }
    } else { nodata }

    $fwRules = Read-JsonSafe (Join-Path $currentAuditPath '09e_firewall_rules_custom.json')
    h3 'カスタムファイアウォールルール'
    if ($fwRules) {
        $sortedRules = @($fwRules) | Sort-Object Direction, DisplayName
        ln '| 表示名 | 方向 | アクション | プロファイル | 有効 | プロトコル |'
        ln '|--------|------|----------|------------|------|----------|'
        foreach ($r in $sortedRules) {
            ln "| $(fesc $r.DisplayName -MaxLen 50) | $(fesc $r.Direction) | $(fesc $r.Action) | $(fesc $r.Profile) | $(fesc $r.Enabled) | $(fesc $r.Protocol) |"
        }
    } else { nodata }

    $defPref = Read-JsonSafe (Join-Path $currentAuditPath '09i_defender_preferences.json')
    h3 'Windows Defender 設定'
    if ($defPref) {
        ln '| 項目 | 値 |'
        ln '|------|-----|'
        ln "| リアルタイム保護 | $(Bool-Str $defPref.DisableRealtimeMonitoring '無効' '有効') |"
        ln "| 動作監視 | $(Bool-Str $defPref.DisableBehaviorMonitoring '無効' '有効') |"
        ln "| 初回検出時ブロック | $(Bool-Str $defPref.DisableBlockAtFirstSeen '無効' '有効') |"
        ln "| IOAVプロテクション | $(Bool-Str $defPref.DisableIOAVProtection '無効' '有効') |"
        ln "| スクリプトスキャン | $(Bool-Str $defPref.DisableScriptScanning '無効' '有効') |"
        ln "| ネットワーク保護 | $(fv $defPref.EnableNetworkProtection) |"
        ln "| PUA保護 | $(fv $defPref.PUAProtection) |"
        ln "| MAPSレポート | $(fv $defPref.MAPSReporting) |"
        ln "| クラウドブロックレベル | $(fv $defPref.CloudBlockLevel) |"
        ln "| スキャンスケジュール | Day=$(fv $defPref.ScanScheduleDay) $(fv $defPref.ScanScheduleTime) |"
        $exPaths = @($defPref.ExclusionPath)  | Where-Object { $_ }
        $exExts  = @($defPref.ExclusionExtension) | Where-Object { $_ }
        $exProcs = @($defPref.ExclusionProcess)   | Where-Object { $_ }
        if ($exPaths.Count -gt 0)  { ln "| 除外パス | $(fesc ($exPaths -join ', ') -MaxLen 80) |" }
        if ($exExts.Count -gt 0)   { ln "| 除外拡張子 | $(fesc ($exExts -join ', ') -MaxLen 80) |" }
        if ($exProcs.Count -gt 0)  { ln "| 除外プロセス | $(fesc ($exProcs -join ', ') -MaxLen 80) |" }
    } else { nodata }

    $certs = Read-JsonSafe (Join-Path $currentAuditPath '18d_certificates.json')
    h3 '証明書 (秘密鍵あり / ユーザーストア)'
    if ($certs) {
        $notable = @($certs) | Where-Object {
            $_.HasPrivateKey -eq $true -or $_.Store -match '\\My$'
        } | Sort-Object Store, Subject
        if ($notable.Count -gt 0) {
            ln '| ストア | サブジェクト | 有効期限 | 秘密鍵 |'
            ln '|--------|------------|---------|--------|'
            foreach ($c in $notable) {
                $storeShort = $c.Store -replace 'Cert:\\', '' -replace 'LocalMachine\\', 'LM\\' -replace 'CurrentUser\\', 'CU\\'
                $notAfter   = ConvertFrom-JsonDate $c.NotAfter
                ln "| $(fesc $storeShort) | $(fesc $c.Subject -MaxLen 60) | $notAfter | $(Bool-Str $c.HasPrivateKey 'あり' 'なし') |"
            }
        } else { ln '(秘密鍵あり証明書なし)' }
    } else { nodata }

    # ═══════════════════════════════════════════════════════
    # 5. ローカルユーザー・グループ
    # ═══════════════════════════════════════════════════════

    h2 '5. ローカルユーザー・グループ'

    $users = Read-JsonSafe (Join-Path $currentAuditPath '10_local_users.json')
    h3 'ローカルユーザー'
    if ($users) {
        $sortedUsers = @($users) | Sort-Object { $_.Name.ToLower() }
        ln '| ユーザー名 | 有効 | 説明 |'
        ln '|-----------|------|------|'
        foreach ($u in $sortedUsers) {
            $enabled = Bool-Str $u.Enabled 'あり' 'なし'
            ln "| $(fesc $u.Name) | $enabled | $(fesc $u.Description '(なし)') |"
        }
    } else { nodata }

    $groups = Read-JsonSafe (Join-Path $currentAuditPath '10b_local_groups.json')
    h3 'Administrators グループ メンバー'
    if ($groups) {
        $adminGrp = @($groups) | Where-Object { $_.Name -eq 'Administrators' } | Select-Object -First 1
        if ($adminGrp -and @($adminGrp.Members).Count -gt 0) {
            ln '| メンバー名 | 種別 |'
            ln '|-----------|------|'
            $sortedMembers = @($adminGrp.Members) | Sort-Object { $_.Name.ToLower() }
            foreach ($m in $sortedMembers) {
                ln "| $(fesc $m.Name) | $(fesc $m.Type) |"
            }
        } else { ln '(メンバーなし)' }
    } else { nodata }

    # ═══════════════════════════════════════════════════════
    # 6. インストール済みアプリ
    # ═══════════════════════════════════════════════════════

    h2 '6. インストール済みアプリ'

    $apps = Read-JsonSafe (Join-Path $currentAuditPath '03_installed_apps_win32.json')
    h3 'Win32 アプリ'
    if ($apps) {
        $sortedApps = @($apps) | Sort-Object { $_.Name.ToLower() }
        ln '| アプリ名 | バージョン | 発行者 |'
        ln '|---------|----------|--------|'
        foreach ($app in $sortedApps) {
            ln "| $(fesc $app.Name) | $(fesc $app.Version) | $(fesc $app.Publisher) |"
        }
    } else { nodata }

    $portableApps = Read-JsonSafe (Join-Path $currentAuditPath '03c_portable_apps_scan.json')
    h3 'ポータブルアプリ (未登録・フォルダ単位)'
    if ($portableApps) {
        $unregistered = @($portableApps) | Where-Object { -not $_.IsRegistered }
        if ($unregistered.Count -gt 0) {
            # フォルダごとに代表1件（ProductNameあり優先）
            $dirSeen = @{}
            $rows    = [System.Collections.Generic.List[object]]::new()
            foreach ($a in ($unregistered | Sort-Object @{Expression={if($_.ProductName){0}else{1}}}, FullPath)) {
                $dir = Split-Path $a.FullPath -Parent
                if (-not $dirSeen.ContainsKey($dir)) { $dirSeen[$dir] = $true; $rows.Add($a) }
            }
            $sorted = $rows | Sort-Object { if($_.ProductName){"0$($_.ProductName)"}else{"1$($_.FullPath)"} }
            ln '| 製品名 | バージョン | メーカー | フォルダ |'
            ln '|--------|----------|---------|---------|'
            foreach ($app in $sorted) {
                $name = if ($app.ProductName) { fesc $app.ProductName } else { fesc (Split-Path $app.FullPath -Leaf) }
                $dir  = fesc (Split-Path $app.FullPath -Parent) -MaxLen 60
                ln "| $name | $(fesc $app.FileVersion) | $(fesc $app.CompanyName -MaxLen 30) | $dir |"
            }
        } else { ln '(ポータブルアプリなし)' }
    } else { nodata }

    $msiApps = Read-JsonSafe (Join-Path $currentAuditPath '03b_installed_apps_msi.json')
    h3 'MSI アプリ'
    if ($msiApps) {
        $sortedMsi = @($msiApps) | Sort-Object { $_.Name.ToLower() }
        ln '| アプリ名 | バージョン | インストール元 |'
        ln '|---------|----------|--------------|'
        foreach ($app in $sortedMsi) {
            $srcShort = fesc $app.Source -MaxLen 50
            ln "| $(fesc $app.Name) | $(fesc $app.Version) | $srcShort |"
        }
    } else { nodata }

    $storeApps = Read-JsonSafe (Join-Path $currentAuditPath '04_installed_apps_store.json')
    h3 'ストアアプリ'
    if ($storeApps) {
        $sortedStore = @($storeApps) | Where-Object { $_.Name -notmatch '^[0-9a-f-]{8,}$' } | Sort-Object { $_.Name.ToLower() }
        ln '| パッケージ名 | バージョン | 発行者 |'
        ln '|-----------|----------|--------|'
        foreach ($app in $sortedStore) {
            $pub = if ($app.Publisher -match 'O=([^,]+)') { $Matches[1].Trim() } else { fesc $app.Publisher -MaxLen 30 }
            ln "| $(fesc $app.Name) | $(fesc $app.Version) | $pub |"
        }
    } else { nodata }

    # ═══════════════════════════════════════════════════════
    # 7. 開発ツール
    # ═══════════════════════════════════════════════════════

    h2 '7. 開発ツール'

    $dev = Read-JsonSafe (Join-Path $currentAuditPath '06_dev_tools.json')
    if ($dev) {
        $script:devRowBuf = [System.Collections.Generic.List[string]]::new()

        function devrow {
            param([string]$Label, [object]$Val)
            if ($null -eq $Val -or "$Val".Trim() -eq '') { return }
            $script:devRowBuf.Add("| $Label | $(fesc $Val) |")
        }

        $ps = $dev.PowerShell
        if ($ps) { $script:devRowBuf.Add("| PowerShell | $(fv $ps.PSVersion) ($(fv $ps.Edition)) |") }

        $dn = $dev.DotNet
        if ($dn -and $dn.Runtimes -and @($dn.Runtimes).Count -gt 0) {
            $runtimes = (@($dn.Runtimes) | ForEach-Object { ($_ -split '\[')[0].Trim() }) -join '; '
            $script:devRowBuf.Add("| .NET ランタイム | $(fesc $runtimes -MaxLen 120) |")
        }

        devrow 'Node.js'         ($dev.Node.node)
        devrow 'npm'             ($dev.Node.npm)
        $pyVer = if ($dev.Python.python) { $dev.Python.python } else { $dev.Python.python3 }
        devrow 'Python'          $pyVer
        devrow 'pip'             ($dev.Python.pip)
        devrow 'Java (java)'     ($dev.Java.java)
        devrow 'Go'              ($dev.Go.go)
        devrow 'Rust (rustc)'    ($dev.Rust.rustc)
        devrow 'Ruby'            ($dev.Ruby.ruby)
        devrow 'Git'             ($dev.Git.version)
        devrow 'Docker'          ($dev.Docker.docker)
        devrow 'curl'            ($dev.OtherTools.curl)
        devrow 'wget'            ($dev.OtherTools.wget)
        devrow 'make'            ($dev.OtherTools.make)
        devrow 'cmake'           ($dev.OtherTools.cmake)
        devrow 'kubectl'         ($dev.OtherTools.kubectl)
        devrow 'helm'            ($dev.OtherTools.helm)
        devrow 'terraform'       ($dev.OtherTools.terraform)
        devrow 'az (Azure CLI)'  ($dev.OtherTools.az)
        devrow 'aws (AWS CLI)'   ($dev.OtherTools.aws)

        if ($script:devRowBuf.Count -gt 0) {
            ln '| ツール | バージョン / 詳細 |'
            ln '|--------|----------------|'
            foreach ($row in $script:devRowBuf) { ln $row }
        } else {
            ln ''
            ln '(インストール済みの開発ツールなし)'
        }
    } else { nodata }

    $psModules = Read-JsonSafe (Join-Path $currentAuditPath '06d_powershell_modules.json')
    h3 'PowerShell モジュール'
    if ($psModules) {
        $sortedMods = @($psModules) | Sort-Object { $_.Name.ToLower() }
        ln '| モジュール名 | バージョン |'
        ln '|------------|----------|'
        foreach ($m in $sortedMods) {
            ln "| $(fesc $m.Name) | $(fesc $m.Version) |"
        }
    } else { nodata }

    # ═══════════════════════════════════════════════════════
    # 8. スタートアップ
    # ═══════════════════════════════════════════════════════

    h2 '8. スタートアップ'

    $startupReg = Read-JsonSafe (Join-Path $currentAuditPath '11_startup_registry.json')
    h3 'レジストリ Run キー'
    if ($startupReg) {
        $sortedReg = @($startupReg) | Sort-Object { $_.Registry }, { $_.Name.ToLower() }
        ln '| レジストリ | 名前 | コマンド |'
        ln '|-----------|------|---------|'
        foreach ($r in $sortedReg) {
            $regShort = $r.Registry `
                -replace 'HKLM:\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion', 'HKLM:\...' `
                -replace 'HKCU:\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion', 'HKCU:\...'
            ln "| $(fesc $regShort) | $(fesc $r.Name) | $(fesc $r.Command -MaxLen 70) |"
        }
    } else { nodata }

    $startupDir = Read-JsonSafe (Join-Path $currentAuditPath '11b_startup_folders.json')
    h3 'スタートアップフォルダ'
    if ($startupDir) {
        $sortedDir = @($startupDir) | Sort-Object { $_.Folder }, { $_.Name.ToLower() }
        ln '| 場所 | ファイル名 |'
        ln '|------|----------|'
        foreach ($s in $sortedDir) {
            $folderShort = $s.Folder `
                -replace [regex]::Escape('C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Startup'), 'AllUsers スタートアップ' `
                -replace ([regex]::Escape('C:\Users\') + '[^\\]+' + [regex]::Escape('\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup')), 'CurrentUser スタートアップ'
            ln "| $(fesc $folderShort) | $(fesc $s.Name) |"
        }
    } else { nodata }

    # ═══════════════════════════════════════════════════════
    # 9. 主要サービス (自動起動)
    # ═══════════════════════════════════════════════════════

    h2 '9. 主要サービス (自動起動)'

    $services = Read-JsonSafe (Join-Path $currentAuditPath '08_services.json')
    if ($services) {
        $autoSvcs = @($services) | Where-Object { $_.StartType -like 'Automatic*' } | Sort-Object Name
        if ($autoSvcs.Count -gt 0) {
            ln '| サービス名 | 表示名 | 実行アカウント |'
            ln '|-----------|--------|-------------|'
            foreach ($svc in $autoSvcs) {
                ln "| $(fesc $svc.Name) | $(fesc $svc.DisplayName) | $(fesc $svc.StartName) |"
            }
        } else { ln '(自動起動サービスなし)' }
    } else { nodata }

    # ═══════════════════════════════════════════════════════
    # 10. GPOスクリプト
    # ═══════════════════════════════════════════════════════

    h2 '10. GPOスクリプト'

    $manifest = Read-JsonSafe (Join-Path $currentAuditPath '20_scripts_manifest.json')
    h3 '収集済みスクリプト'
    if ($manifest) {
        $sortedManifest = @($manifest) | Sort-Object Category
        ln '| カテゴリ | ファイル名 | サイズ (B) | 最終更新 |'
        ln '|---------|----------|-----------|---------|'
        foreach ($m in $sortedManifest) {
            ln "| $(fesc $m.Category) | $(fesc $m.CopiedAs) | $(fv $m.SizeBytes) | $(fesc $m.LastWrite) |"
        }
    } else { nodata }

    $gpoLogon = Read-JsonSafe (Join-Path $currentAuditPath '10i_gpo_logon_scripts.json')
    h3 'ログオン/ログオフスクリプト (GPO レジストリ)'
    if ($gpoLogon) {
        $items = @($gpoLogon) | Where-Object { $_.ScriptPath }
        if ($items.Count -gt 0) {
            ln '| 種別 | スクリプト |'
            ln '|------|----------|'
            foreach ($g in $items) { ln "| $(fesc $g.Type) | $(fesc $g.ScriptPath) |" }
        } else { ln '(ログオン/ログオフスクリプトなし)' }
    } else {
        ln ''
        ln '(データなし - レジストリキー未設定)'
    }

    $folderRedir = Read-JsonSafe (Join-Path $currentAuditPath '10i3_folder_redirection.json')
    h3 'フォルダリダイレクト (シェルフォルダ)'
    if ($folderRedir) {
        ln '| フォルダ名 | パス |'
        ln '|----------|------|'
        foreach ($groupProp in $folderRedir.PSObject.Properties) {
            foreach ($fp in $groupProp.Value.PSObject.Properties) {
                ln "| $(fesc $fp.Name) | $(fesc $fp.Value) |"
            }
        }
    } else { nodata }

    # ═══════════════════════════════════════════════════════
    # 11. Windows オプション機能 (有効のみ)
    # ═══════════════════════════════════════════════════════

    h2 '11. Windows オプション機能 (有効のみ)'

    $features = Read-JsonSafe (Join-Path $currentAuditPath '07_windows_optional_features.json')
    h3 'オプション機能'
    if ($features) {
        $enabled = @($features) | Where-Object { $_.State -eq 'Enabled' } | Sort-Object Feature
        if ($enabled.Count -gt 0) {
            ln '| 機能名 |'
            ln '|--------|'
            foreach ($f in $enabled) { ln "| $(fesc $f.Feature) |" }
        } else { ln '(有効なオプション機能はありません)' }
    } else { nodata }

    $caps = Read-JsonSafe (Join-Path $currentAuditPath '07b_windows_capabilities.json')
    h3 'Windows 機能 (Capabilities・インストール済み)'
    if ($caps) {
        $instCaps = @($caps) | Where-Object { $_.State -eq 'Installed' } | Sort-Object Name
        if ($instCaps.Count -gt 0) {
            ln '| 機能名 |'
            ln '|--------|'
            foreach ($c in $instCaps) { ln "| $(fesc $c.Name) |" }
        } else { ln '(インストール済みCapabilityはありません)' }
    } else { nodata }

    # ═══════════════════════════════════════════════════════
    # 12. スケジュールタスク (非Microsoft)
    # ═══════════════════════════════════════════════════════

    h2 '12. スケジュールタスク (非Microsoft)'

    $tasks = Read-JsonSafe (Join-Path $currentAuditPath '11c_scheduled_tasks.json')
    if ($tasks) {
        $sortedTasks = @($tasks) | Sort-Object { $_.TaskName.ToLower() }
        if ($sortedTasks.Count -gt 0) {
            ln '| タスク名 | フォルダ | アクション |'
            ln '|---------|---------|----------|'
            foreach ($t in $sortedTasks) {
                $action  = if ($t.Actions -and @($t.Actions).Count -gt 0) {
                               fesc @($t.Actions)[0] -MaxLen 60
                           } else { '-' }
                ln "| $(fesc $t.TaskName) | $(fesc $t.TaskPath) | $action |"
            }
        } else { ln '(非Microsoftタスクなし)' }
    } else { nodata }

    # ═══════════════════════════════════════════════════════
    # 13. シェル設定
    # ═══════════════════════════════════════════════════════

    h2 '13. シェル設定'

    $shellSet   = Read-JsonSafe (Join-Path $currentAuditPath '18c_shell_settings.json')
    $profPolicy = Read-JsonSafe (Join-Path $currentAuditPath '10i6_profile_policy.json')

    h3 'シェル・エクスプローラー'
    if ($shellSet) {
        ln '| 項目 | 値 |'
        ln '|------|-----|'
        ln "| デフォルトブラウザ | $(fesc $shellSet.DefaultBrowserProgId) |"
        ln "| デフォルトシェル | $(fesc $shellSet.DefaultShell) |"
        if ($shellSet.FileExplorer) {
            $fe = $shellSet.FileExplorer
            ln "| 隠しファイルを表示 | $(Bool-Str $fe.ShowHiddenFiles '表示' '非表示') |"
            ln "| 拡張子を表示 | $(Bool-Str $fe.ShowFileExtensions '表示' '非表示') |"
            ln "| フルパスを表示 | $(Bool-Str $fe.ShowFullPath '表示' '非表示') |"
        }
    } else { nodata }

    h3 'Winlogon (ログオンシェル)'
    if ($profPolicy) {
        $wlProp = $profPolicy.PSObject.Properties | Where-Object { $_.Name -match 'Winlogon' } | Select-Object -First 1
        if ($wlProp) {
            ln '| 項目 | 値 |'
            ln '|------|-----|'
            foreach ($vp in $wlProp.Value.PSObject.Properties) {
                ln "| $(fesc $vp.Name) | $(fesc $vp.Value) |"
            }
        } else { ln '(Winlogon エントリなし)' }
    } else { nodata }

    $wtSettings = Read-JsonSafe (Join-Path $currentAuditPath '18b_windows_terminal_settings.json')
    h3 'Windows Terminal プロファイル'
    if ($wtSettings -and $wtSettings.profiles -and $wtSettings.profiles.list) {
        $defGuid = fv $wtSettings.defaultProfile
        ln '| プロファイル名 | コマンドライン | 既定 |'
        ln '|-------------|-------------|------|'
        foreach ($p in @($wtSettings.profiles.list)) {
            $isDefault = if ($p.guid -eq $defGuid) { '★' } else { '' }
            ln "| $(fesc $p.name) | $(fesc $p.commandline) | $isDefault |"
        }
    } else { nodata }

    # ═══════════════════════════════════════════════════════
    # 14. レジストリポリシー
    # ═══════════════════════════════════════════════════════

    h2 '14. レジストリポリシー'

    $regPol = Read-JsonSafe (Join-Path $currentAuditPath '19c_registry_policies.json')

    h3 'コンピューターポリシー (HKLM\Policies)'
    if ($regPol -and $regPol.ComputerPolicies -and @($regPol.ComputerPolicies).Count -gt 0) {
        ln '| キーパス | 名前 | 値 |'
        ln '|---------|------|-----|'
        foreach ($p in @($regPol.ComputerPolicies)) {
            $ks = $p.KeyPath.Replace('HKEY_LOCAL_MACHINE\SOFTWARE\Policies\', '')
            ln "| $(fesc $ks -MaxLen 60) | $(fesc $p.Name -MaxLen 50) | $(fesc $p.Value -MaxLen 60) |"
        }
    } else { nodata }

    h3 'コンピューターポリシー Legacy (HKLM\...\Policies)'
    if ($regPol -and $regPol.ComputerPoliciesLegacy -and @($regPol.ComputerPoliciesLegacy).Count -gt 0) {
        ln '| キーパス | 名前 | 値 |'
        ln '|---------|------|-----|'
        foreach ($p in @($regPol.ComputerPoliciesLegacy)) {
            $ks = $p.KeyPath.Replace('HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\', '')
            ln "| $(fesc $ks -MaxLen 60) | $(fesc $p.Name -MaxLen 50) | $(fesc $p.Value -MaxLen 60) |"
        }
    } else { nodata }

    h3 'ユーザーポリシー'
    if ($regPol -and $regPol.UserPoliciesPerUser) {
        $hasAnyUser = $false
        foreach ($userProp in $regPol.UserPoliciesPerUser.PSObject.Properties) {
            $uData = $userProp.Value
            $allP  = @()
            if ($uData.UserPolicies)       { $allP += @($uData.UserPolicies) }
            if ($uData.UserPoliciesLegacy) { $allP += @($uData.UserPoliciesLegacy) }
            if ($allP.Count -eq 0) { continue }
            $hasAnyUser = $true
            ln ''
            ln "**$($userProp.Name)**"
            ln ''
            ln '| キーパス | 名前 | 値 |'
            ln '|---------|------|-----|'
            foreach ($p in $allP) {
                $ks = $p.KeyPath -replace 'HKEY_USERS\\S-[0-9-]+\\SOFTWARE\\Policies\\', ''
                $ks = $ks         -replace 'HKEY_USERS\\S-[0-9-]+\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Policies\\', ''
                ln "| $(fesc $ks -MaxLen 60) | $(fesc $p.Name -MaxLen 50) | $(fesc $p.Value -MaxLen 60) |"
            }
        }
        if (-not $hasAnyUser) { ln '(なし)' }
    } else { nodata }

    # ═══════════════════════════════════════════════════════
    # 15. 環境変数
    # ═══════════════════════════════════════════════════════

    h2 '15. 環境変数'

    $envVars = Read-JsonSafe (Join-Path $currentAuditPath '10c_environment_variables.json')

    h3 'システム環境変数'
    if ($envVars -and $envVars.System) {
        ln '| 変数名 | 値 |'
        ln '|--------|-----|'
        foreach ($v in @($envVars.System) | Sort-Object Name) {
            ln "| $(fesc $v.Name) | $(fesc $v.Value -MaxLen 100) |"
        }
    } else { nodata }

    h3 'ユーザー環境変数'
    if ($envVars -and $envVars.User) {
        ln '| 変数名 | 値 |'
        ln '|--------|-----|'
        foreach ($v in @($envVars.User) | Sort-Object Name) {
            ln "| $(fesc $v.Name) | $(fesc $v.Value -MaxLen 100) |"
        }
    } else { nodata }

    # ═══════════════════════════════════════════════════════
    # 16. ロケール・言語
    # ═══════════════════════════════════════════════════════

    h2 '16. ロケール・言語'

    $locale = Read-JsonSafe (Join-Path $currentAuditPath '13_locale_language.json')
    if ($locale) {
        ln '| 項目 | 値 |'
        ln '|------|-----|'
        ln "| タイムゾーン | $(fesc $locale.TimeZoneDisplay) |"
        ln "| カルチャ | $(fesc $locale.Culture) |"
        ln "| UIカルチャ | $(fesc $locale.UICulture) |"
        ln "| システムロケール | $(fesc $locale.SystemLocale) |"
        ln "| 日付フォーマット | $(fesc $locale.DateFormat) |"
        ln "| 時刻フォーマット | $(fesc $locale.TimeFormat) |"
        ln "| 週の開始日 | $(fesc $locale.FirstDayOfWeek) |"
        ln "| 所在地 | $(fesc $locale.HomeLocation) |"
        if ($locale.InputLanguages -and @($locale.InputLanguages).Count -gt 0) {
            ln ''
            ln '**入力言語**'
            ln ''
            ln '| 言語 | タグ |'
            ln '|------|------|'
            foreach ($lang in @($locale.InputLanguages)) {
                ln "| $(fesc $lang.Autonym) | $(fesc $lang.LanguageTag) |"
            }
        }
    } else { nodata }

    # ═══════════════════════════════════════════════════════
    # 17. プリンター
    # ═══════════════════════════════════════════════════════

    h2 '17. プリンター'

    $printers = Read-JsonSafe (Join-Path $currentAuditPath '17b_printers.json')
    if ($printers) {
        ln '| プリンター名 | ドライバー | ポート | 共有 |'
        ln '|-----------|----------|--------|------|'
        foreach ($p in @($printers)) {
            $shared = Bool-Str $p.Shared '共有' '非共有'
            ln "| $(fesc $p.Name) | $(fesc $p.DriverName) | $(fesc $p.PortName) | $shared |"
        }
    } else { nodata }

    # ═══════════════════════════════════════════════════════
    # 18. デバイスドライバー
    # ═══════════════════════════════════════════════════════

    h2 '18. デバイスドライバー'

    $drivers = Read-JsonSafe (Join-Path $currentAuditPath '16_drivers.json')
    if ($drivers) {
        $validDrivers = @($drivers) | Where-Object { $_.DeviceName } | Sort-Object DeviceName
        if ($validDrivers.Count -gt 0) {
            ln '| デバイス名 | ドライバーバージョン | メーカー | 署名 | INF |'
            ln '|-----------|------------------|---------|------|-----|'
            foreach ($d in $validDrivers) {
                $signed = Bool-Str $d.IsSigned '済' 'なし'
                ln "| $(fesc $d.DeviceName -MaxLen 50) | $(fesc $d.DriverVersion) | $(fesc $d.Manufacturer -MaxLen 30) | $signed | $(fesc $d.InfName) |"
            }
        } else { ln '(デバイス名あり ドライバーなし)' }
    } else { nodata }

    # ═══════════════════════════════════════════════════════
    # 19. フォント
    # ═══════════════════════════════════════════════════════

    h2 '19. フォント'

    $fonts = Read-JsonSafe (Join-Path $currentAuditPath '17_fonts.json')
    if ($fonts) {
        $sortedFonts = @($fonts) | Sort-Object Name
        ln '| フォント名 | 種別 | サイズ (KB) |'
        ln '|-----------|------|-----------|'
        foreach ($f in $sortedFonts) {
            ln "| $(fesc $f.Name) | $(fesc $f.Extension) | $(fv $f.SizeKB) |"
        }
    } else { nodata }

    # ─────────────────────────────────────────────────
    # ファイル書き込み (UTF-8 BOMなし)
    # ─────────────────────────────────────────────────

    $content   = $lines -join "`n"
    $utf8NoBom = New-Object System.Text.UTF8Encoding $false
    [System.IO.File]::WriteAllText($outFile, $content, $utf8NoBom)

    Write-Host "生成: $outFile" -ForegroundColor Green

} # end foreach
