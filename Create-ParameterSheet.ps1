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

function Normalize-InstallDate {
    param([string]$d)
    if ([string]::IsNullOrEmpty($d) -or $d -eq '-') { return '-' }
    if ($d -match '^\d{8}$') {
        return "$($d.Substring(0,4))-$($d.Substring(4,2))-$($d.Substring(6,2))"
    }
    if ($d -match '^(\d{1,2})/(\d{1,2})/(\d{4})$') {
        return '{0}-{1:D2}-{2:D2}' -f $Matches[3], [int]$Matches[1], [int]$Matches[2]
    }
    return $d
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
        ln "| インストール日 | $(ConvertFrom-JsonDate $os.InstallDate) |"
        ln "| 最終起動 | $(ConvertFrom-JsonDate $os.LastBootUpTime) |"

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
        ln '| # | モデル | 容量 (GB) | 接続 | タイプ | 状態 |'
        ln '|---|--------|-----------|------|--------|------|'
        foreach ($d in @($disks)) {
            ln "| $(fv $d.DiskNumber) | $(fesc $d.FriendlyName) | $(fv $d.SizeGB) | $(fesc $d.BusType) | $(fesc $d.PartitionStyle) | $(fesc $d.OperationalStatus) |"
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
        ln '| アダプター名 | 状態 | MAC | IPアドレス | CIDR | ゲートウェイ | DNS | 速度 |'
        ln '|-------------|------|-----|-----------|------|------------|-----|------|'
        foreach ($a in $adapters) {
            $ipStr  = if ($a.IPv4Address)      { (@($a.IPv4Address) | Where-Object { $_ }) -join ', ' } else { '-' }
            $cidr   = if ($a.IPv4PrefixLength) { '/' + (@($a.IPv4PrefixLength)[0]) } else { '-' }
            $gwStr  = if ($a.DefaultGateway)   { (@($a.DefaultGateway) | Where-Object { $_ }) -join ', ' } else { '-' }
            $dnsStr = if ($a.DNSServers)       { (@($a.DNSServers) | Where-Object { $_ }) -join ', ' } else { '-' }
            ln "| $(fesc $a.Name) | $(fesc $a.Status) | $(fesc $a.MacAddress) | $ipStr | $cidr | $(fesc $gwStr) | $(fesc $dnsStr) | $(fesc $a.LinkSpeedMbps) |"
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

    # ═══════════════════════════════════════════════════════
    # 5. ローカルユーザー・グループ
    # ═══════════════════════════════════════════════════════

    h2 '5. ローカルユーザー・グループ'

    $users = Read-JsonSafe (Join-Path $currentAuditPath '10_local_users.json')
    h3 'ローカルユーザー'
    if ($users) {
        $sortedUsers = @($users) | Sort-Object { $_.Name.ToLower() }
        ln '| ユーザー名 | 有効 | 最終ログオン | 説明 |'
        ln '|-----------|------|------------|------|'
        foreach ($u in $sortedUsers) {
            $logon   = ConvertFrom-JsonDate $u.LastLogon
            $enabled = Bool-Str $u.Enabled 'あり' 'なし'
            ln "| $(fesc $u.Name) | $enabled | $logon | $(fesc $u.Description '(なし)') |"
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
    # 6. インストール済みアプリ (Win32)
    # ═══════════════════════════════════════════════════════

    h2 '6. インストール済みアプリ (Win32)'

    $apps = Read-JsonSafe (Join-Path $currentAuditPath '03_installed_apps_win32.json')
    if ($apps) {
        $sortedApps = @($apps) | Sort-Object { $_.Name.ToLower() }
        ln '| アプリ名 | バージョン | 発行者 | インストール日 |'
        ln '|---------|----------|--------|--------------|'
        foreach ($app in $sortedApps) {
            $idate = Normalize-InstallDate (fv $app.InstallDate)
            ln "| $(fesc $app.Name) | $(fesc $app.Version) | $(fesc $app.Publisher) | $idate |"
        }
    } else { nodata }

    # ═══════════════════════════════════════════════════════
    # 7. 開発ツール
    # ═══════════════════════════════════════════════════════

    h2 '7. 開発ツール'

    $dev = Read-JsonSafe (Join-Path $currentAuditPath '06_dev_tools.json')
    if ($dev) {
        ln '| ツール | バージョン / 詳細 |'
        ln '|--------|----------------|'

        function devrow {
            param([string]$Label, [object]$Val)
            $disp = if ($null -eq $Val) { '-' }
                    elseif ("$Val".Trim() -eq '') { '(インストール済み)' }
                    else { fesc $Val }
            $script:lines.Add("| $Label | $disp |")
        }

        $ps = $dev.PowerShell
        if ($ps) { ln "| PowerShell | $(fv $ps.PSVersion) ($(fv $ps.Edition)) |" }

        $dn = $dev.DotNet
        if ($dn -and $dn.Runtimes -and @($dn.Runtimes).Count -gt 0) {
            $runtimes = (@($dn.Runtimes) | ForEach-Object { ($_ -split '\[')[0].Trim() }) -join '; '
            ln "| .NET ランタイム | $(fesc $runtimes -MaxLen 120) |"
        } else {
            ln '| .NET ランタイム | - |'
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
            ln '| サービス名 | 表示名 | 状態 | 実行アカウント |'
            ln '|-----------|--------|------|-------------|'
            foreach ($svc in $autoSvcs) {
                $status = fesc $svc.Status
                if ($svc.Status -eq 'Stopped') { $status = "$status (停止中)" }
                ln "| $(fesc $svc.Name) | $(fesc $svc.DisplayName) | $status | $(fesc $svc.StartName) |"
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

    # ═══════════════════════════════════════════════════════
    # 11. Windows オプション機能 (有効のみ)
    # ═══════════════════════════════════════════════════════

    h2 '11. Windows オプション機能 (有効のみ)'

    $features = Read-JsonSafe (Join-Path $currentAuditPath '07_windows_optional_features.json')
    if ($features) {
        $enabled = @($features) | Where-Object { $_.State -eq 'Enabled' } | Sort-Object Feature
        if ($enabled.Count -gt 0) {
            ln '| 機能名 |'
            ln '|--------|'
            foreach ($f in $enabled) { ln "| $(fesc $f.Feature) |" }
        } else { ln '(有効なオプション機能はありません)' }
    } else { nodata }

    # ═══════════════════════════════════════════════════════
    # 12. スケジュールタスク (非Microsoft)
    # ═══════════════════════════════════════════════════════

    h2 '12. スケジュールタスク (非Microsoft)'

    $tasks = Read-JsonSafe (Join-Path $currentAuditPath '11c_scheduled_tasks.json')
    if ($tasks) {
        $sortedTasks = @($tasks) | Sort-Object { $_.TaskName.ToLower() }
        if ($sortedTasks.Count -gt 0) {
            ln '| タスク名 | フォルダ | 状態 | 最終実行 | アクション |'
            ln '|---------|---------|------|---------|----------|'
            foreach ($t in $sortedTasks) {
                $lastRun = ConvertFrom-JsonDate $t.LastRunTime
                $action  = if ($t.Actions -and @($t.Actions).Count -gt 0) {
                               fesc @($t.Actions)[0] -MaxLen 60
                           } else { '-' }
                ln "| $(fesc $t.TaskName) | $(fesc $t.TaskPath) | $(fesc $t.State) | $lastRun | $action |"
            }
        } else { ln '(非Microsoftタスクなし)' }
    } else { nodata }

    # ─────────────────────────────────────────────────
    # ファイル書き込み (UTF-8 BOMなし)
    # ─────────────────────────────────────────────────

    $content   = $lines -join "`n"
    $utf8NoBom = New-Object System.Text.UTF8Encoding $false
    [System.IO.File]::WriteAllText($outFile, $content, $utf8NoBom)

    Write-Host "生成: $outFile" -ForegroundColor Green

} # end foreach
