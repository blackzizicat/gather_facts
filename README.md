# Windows 環境調査ツール - 調査項目リスト

## 目的
Windows クライアント・サーバー環境の設定・構成情報を網羅的に収集する。
組織管理下のマシンにおける制御・制限設定の把握も目的に含む。
Windows Server では、インストールされている役割（AD DS・Hyper-V・IIS・DNS・DHCP・FS・RDS）に応じて
専用の情報を追加収集する（S01〜S08 ステップ）。

---

## 調査カテゴリ一覧

### 1. システム基本情報
- OS バージョン・ビルド番号・エディション
- コンピュータ名・ワークグループ/ドメイン
- ハードウェア構成（CPU、RAM、GPU、マザーボード）
- BIOS/UEFI 情報（バージョン、セキュアブート設定）
- インストール日時

### 2. ディスク・ストレージ
- **物理ディスク一覧**（内蔵・外付けHDD/SSD・USBストレージ含む）
  - ディスクモデル・容量・インターフェース種別（SATA/NVMe/USB等）・シリアル番号
  - `Get-Disk` 優先、失敗時は WMI (`Win32_DiskDrive`) フォールバック
  - フォールバック時も WMI アソシエーションでパーティション・ボリューム情報を取得
- **パーティション情報**（ドライブレター・ファイルシステム・サイズ・空き容量）
- **論理ドライブ一覧**（全ドライブレター・DriveType・容量・空き容量・使用率）

### 3. インストール済みソフトウェア
- Win32 アプリ（Add/Remove Programs 相当）
- MSI インストーラー直接インストール（ProductCode・UpgradeCode 付き詳細）
- ポータブルアプリ（インストーラーなし・EXE 直置き）のスキャン
  - スキャン対象: Program Files / Program Files (x86) / AppData\Local\Programs / デスクトップ / ダウンロード / ドキュメント / PortableApps
  - Win32 登録済みアプリとの照合結果（`IsRegistered` フラグ）を付与
- Microsoft Store アプリ (AppX/MSIX)
- winget パッケージ一覧（プログレスバー表示行を除去して記録）
- Chocolatey パッケージ一覧（インストール済みの場合）
- Scoop パッケージ一覧（インストール済みの場合）

### 4. 開発ツール・ランタイム
- .NET ランタイム・SDK バージョン
- Node.js / npm バージョン・グローバルパッケージ
- Python バージョン・pip パッケージ
- Java / JDK バージョン
- Go / Rust / Ruby バージョン
- Git 設定（グローバル config）
- Docker 設定
- Visual Studio Code 拡張機能

### 5. Windows 機能・オプション機能
- Windows オプション機能（Hyper-V、IIS 等）
- Windows Capabilities
- PowerShell インストール済みモジュール

### 6. サービス
- 実行中・停止中のサービス一覧
- スタートアップの種類（自動/手動/無効）
- サービスの実行パス・実行アカウント

### 7. ネットワーク設定
- ネットワークアダプター一覧・設定（IP/サブネット/ゲートウェイ/DNS/IPv6）
- Wi-Fi プロファイル（SSID 一覧、パスワード除く）
- hosts ファイル
- プロキシ設定
- **ファイアウォールプロファイル**（ドメイン/プライベート/パブリック各プロファイルのON/OFF・既定動作）
- ファイアウォールカスタムルール（インバウンド/アウトバウンド）

### 8. セキュリティソフト・防御設定
- **登録済みセキュリティ製品**（Windows SecurityCenter2 WMI）
  - サードパーティ製 AV・ファイアウォール・スパイウェア対策ソフトの有無・有効状態・定義最新状態
- **Windows Defender / Microsoft Defender Antivirus ステータス**（`Get-MpComputerStatus`）
  - リアルタイム保護・動作監視・Tamper Protection 等の有効/無効
  - エンジンバージョン・定義バージョン・最終更新日
  - 脅威数・検疫数・フルスキャン/クイックスキャン実施日時
- **Windows Defender 設定**（`Get-MpPreference`）
  - 除外設定（パス・拡張子・プロセス・IPアドレス）
  - Attack Surface Reduction (ASR) ルール一覧
  - クラウド保護（MAPS）レベル・サンプル送信設定
  - ネットワーク保護・PUA 保護設定
  - スキャンスケジュール
- **Microsoft Defender for Endpoint (MDE) オンボーディング状態**
  - `Sense`/`MsSense` サービスの稼働状態
  - オンボーディング状態（OnboardingState）・組織ID（OrgId）・最終接続時刻

### 9. ユーザー・セキュリティ設定
- ローカルユーザーアカウント一覧（SID・有効状態・最終ログオン等）
- ローカルグループ・メンバーシップ
- 環境変数（ユーザー/システム）
- UAC 設定レベル
- BitLocker 状態（管理者権限時）

### 10. 組織管理・ユーザー初期化設定
- 全ユーザープロファイル一覧（ローカル/ローミング/必須プロファイル検出）
- 自動ログイン設定（Winlogon: AutoAdminLogon・Shell・Userinit）
- ユーザーごとの詳細（デスクトップファイル・スタートアップ・Run レジストリ）
- **GPO ログオン・ログオフ・起動・シャットダウンスクリプト（コンピューター側）**
- **ユーザー側 GPO ログオン・ログオフスクリプト（HKCU 側）**
- **フォルダリダイレクト設定**（デスクトップ・ドキュメント等のネットワーク転送先）
- **デフォルトユーザープロファイルの状態**（NTUSER.DAT 更新日・デスクトップ内容・Run レジストリ）
- **ドメインアカウントのログオンスクリプトパス**（`net user` 経由）
- **プロファイルの上書きポリシー**（DeleteCachedCopies・CleanupProfiles 等）
- UWF (Unified Write Filter) 状態（キオスク・共用PC での毎回リセット設定）
- 自動ログイン設定（GPO ログオン・ログオフスクリプト経由）

### 11. スタートアップ・スケジュールタスク
- スタートアップアプリ（レジストリ Run/RunOnce・スタートアップフォルダ）
- タスクスケジューラ登録タスク（非 Microsoft タスクのみ）

### 12. 電源・パフォーマンス設定
- 電源プラン（アクティブ・全プラン設定詳細）
- 仮想メモリ（ページファイル）設定

### 13. 地域・言語・入力設定
- タイムゾーン
- 地域・ロケール設定
- キーボードレイアウト・入力メソッド（IME）

### 14. ディスプレイ設定
- 解像度・リフレッシュレート・ビット深度
- DPI スケーリング

### 15. ファイルシステム・共有設定
- 共有フォルダ一覧（隠し共有除く）
- **ネットワークドライブ一覧**（4手段で収集）
  - `Get-PSDrive`（現在マウント中）
  - `Win32_MappedLogicalDisk`（WMI）
  - `HKCU:\Network`（永続マップのレジストリ）
  - `net use` コマンド

### 16. ドライバー
- インストール済みデバイスドライバー一覧（バージョン・署名状態・製造元）
- **デバイス ハードウェア ID 一覧**（`Win32_PnPEntity` から HardwareID / CompatibleID を取得）
- **専門周辺機器の詳細情報**（キーワード・USB VID フィルタで絞り込み、ドライバー情報を結合）
  - 対象カテゴリ: オーディオ機器・映像機器・デジタイザー等の専門USBデバイス（Focusrite / MOTU / Elgato / Blackmagic / AVerMedia / Wacom / Huion 等）
  - デバイス情報・ドライバー情報の結合
  - ベンダー固有レジストリキー（型番・設定情報）

### 17. フォント・プリンター
- インストール済みフォント一覧
- インストール済みプリンター（ドライバー・ポート・既定設定）

### 18. 証明書
- ローカルマシン・現在ユーザーの証明書ストア（My・Root）

### 19. シェル・ターミナル設定
- PowerShell プロファイル内容（全スコープ）
- Windows Terminal settings.json
- シェル・エクスプローラー設定（デフォルトブラウザ・隠しファイル表示等）

### 20. ポリシー設定
- GPO 適用結果テキスト（`gpresult /r`、Unicode出力で文字化け防止）
- GPO 適用結果 HTML 詳細レポート
  - **管理者権限あり**: `Win32_UserProfile` でプロファイルを列挙し、`gpresult /user [ユーザー名] /h` をユーザーごとに実行して全ユーザー分を収集
  - **管理者権限なし**: 実行ユーザー分のみ（`gpresult /h`）
- レジストリポリシーキー値（`HKLM:\SOFTWARE\Policies\`・`...\CurrentVersion\Policies\`）
  - **管理者権限あり**: 全ユーザーの `HKCU` ポリシーを収集。ログイン中ユーザーは `HKU\<SID>` を直接スキャン、ログアウト済みユーザーは `NTUSER.DAT` を `reg load` でマウントしてスキャン後にアンロード
  - **管理者権限なし**: 実行ユーザー分のみ
- 監査ポリシー（`auditpol /get /category:*`）
- パスワード・アカウントロックアウトポリシー（`net accounts`）
- ローカルセキュリティポリシー全体（`secedit /export`、管理者権限時のみ）

### 21. システム情報（msinfo32 システムの要約）
- `msinfo32.exe /report` を使用してシステムの要約を取得
- OS名・バージョン・プロセッサ・物理メモリ・BIOS バージョン・システム製造元など
- 日本語・英語環境どちらにも対応

### 20. スクリプト・バッチファイルの実体収集
- GPO スクリプトディレクトリのスキャン（Machine/User の Startup・Shutdown・Logon・Logoff）
- スタートアップフォルダ内のスクリプトファイル
- GPO スクリプトレジストリ参照先ファイル
- タスクスケジューラのアクション・引数が参照するスクリプトファイル
- 収集対象拡張子: `.bat` `.cmd` `.ps1` `.vbs` `.wsf` `.js` `.adm` `.admx`
- コピー先: `20_scripts/` ディレクトリ（カテゴリプレフィックス付きで格納）

---

## 出力ファイル一覧

| ファイル名 | 内容 |
|-----------|------|
| `00_INDEX.md` | 収集結果のインデックス |
| `00_REPORT.html` | 収集結果の HTML レポート（`-GenerateHtml` 指定時のみ生成） |
| `01_system_info.json` | OS・CPU・GPU・BIOS・マザーボード情報 |
| `02_disk_partitions.json` | 物理ディスク・パーティション情報 |
| `02b_logical_disks.json` | 論理ドライブ一覧（容量・空き容量・使用率） |
| `03_installed_apps_win32.json` | Win32 アプリ一覧 |
| `03b_installed_apps_msi.json` | MSI インストールアプリ詳細 |
| `03c_portable_apps_scan.json` | ポータブルアプリスキャン結果 |
| `04_installed_apps_store.json` | Store アプリ (AppX/MSIX) |
| `05_package_managers.txt` | winget / Chocolatey / Scoop パッケージ一覧 |
| `06_dev_tools.json` | 開発ツール・ランタイムバージョン |
| `06b_vscode_extensions.txt` | VSCode 拡張機能一覧 |
| `06d_powershell_modules.json` | PowerShell インストール済みモジュール |
| `07_windows_optional_features.json` | Windows オプション機能 |
| `07b_windows_capabilities.json` | Windows Capabilities |
| `08_services.json` | サービス一覧 |
| `09_network_adapters.json` | ネットワークアダプター・IP設定 |
| `09b_hosts_file.txt` | hosts ファイル |
| `09c_proxy_settings.json` | プロキシ設定 |
| `09d_wifi_profiles.txt` | Wi-Fi プロファイル |
| `09e_firewall_rules_custom.json` | ファイアウォールカスタムルール |
| `09f_firewall_profiles.json` | ファイアウォールプロファイル（ドメイン/プライベート/パブリック） |
| `09g_security_products.json` | 登録済みセキュリティ製品（AV・FW・AS） |
| `09h_defender_status.json` | Windows Defender ステータス |
| `09i_defender_preferences.json` | Windows Defender 設定・除外・ASRルール |
| `09j_defender_endpoint.json` | Microsoft Defender for Endpoint 状態 |
| `10_local_users.json` | ローカルユーザー一覧 |
| `10b_local_groups.json` | ローカルグループ・メンバー |
| `10c_environment_variables.json` | 環境変数 |
| `10d_uac_settings.json` | UAC 設定 |
| `10e_bitlocker.json` | BitLocker 状態 |
| `10f_user_profiles.json` | 全ユーザープロファイル一覧 |
| `10g_autologin_settings.json` | 自動ログイン・Winlogon 設定 |
| `10h_per_user_details.json` | ユーザーごとのデスクトップ・スタートアップ・Run |
| `10i_gpo_logon_scripts.json` | GPO ログオン/ログオフスクリプト（コンピューター側） |
| `10i2_user_gpo_logon_scripts.json` | GPO ログオン/ログオフスクリプト（ユーザー側 HKCU） |
| `10i3_folder_redirection.json` | フォルダリダイレクト設定 |
| `10i4_default_user_profile.json` | デフォルトユーザープロファイル状態 |
| `10i5_domain_logon_script.json` | ドメインアカウントのログオンスクリプトパス |
| `10i6_profile_policy.json` | プロファイルの上書きポリシー |
| `10j_uwf_status.txt` | UWF (Unified Write Filter) 状態 |
| `11_startup_registry.json` | スタートアップ（レジストリ Run/RunOnce） |
| `11b_startup_folders.json` | スタートアップフォルダ |
| `11c_scheduled_tasks.json` | スケジュールタスク（非 Microsoft） |
| `12_power_settings.txt` | 電源プラン設定 |
| `12b_virtual_memory.json` | 仮想メモリ設定 |
| `13_locale_language.json` | 地域・言語・入力設定 |
| `14_display_settings.json` | ディスプレイ設定・DPI |
| `15_shared_folders.json` | 共有フォルダ一覧 |
| `15b_network_drives.json` | ネットワークドライブ一覧 |
| `16_drivers.json` | デバイスドライバー一覧 |
| `16b_device_hardware_ids.json` | デバイス ハードウェア ID 一覧（Win32_PnPEntity） |
| `16c_specialized_devices.json` | 専門周辺機器詳細（オーディオ・映像・デジタイザー等の専門USBデバイス） |
| `17_fonts.json` | フォント一覧 |
| `17b_printers.json` | プリンター一覧 |
| `18_powershell_profiles.txt` | PowerShell プロファイル内容 |
| `18b_windows_terminal_settings.json` | Windows Terminal 設定 |
| `18c_shell_settings.json` | シェル・エクスプローラー設定 |
| `18d_certificates.json` | 証明書一覧 |
| `19_gpresult.txt` | GPO 適用結果（テキスト、実行ユーザー分） |
| `19b_gpresult_[ユーザー名].html` | GPO 適用結果（HTML 詳細、管理者時はユーザーごとに生成） |
| `19c_registry_policies.json` | レジストリポリシーキー値（管理者時は全ユーザーの HKCU を含む） |
| `19d_audit_policy.txt` | 監査ポリシー |
| `19e_account_policy.txt` | アカウント・ロックアウトポリシー |
| `19f_local_security_policy.inf` | ローカルセキュリティポリシー全体 |
| `20_scripts/` | 収集したスクリプト・バッチファイル格納ディレクトリ |
| `20_scripts_manifest.json` | スクリプト・バッチファイル収集マニフェスト（収集元パス・カテゴリ・サイズ等） |
| `21_msinfo32_summary.json` | msinfo32 システムの要約（OS・プロセッサ・メモリ・BIOS 等） |
| `S01_server_roles.json` | サーバー役割・機能一覧（Windows Server & 管理者権限時のみ） |
| `S02_active_directory.json` | Active Directory 情報（AD DS 役割インストール時のみ） |
| `S03_hyperv.json` | Hyper-V 情報（Hyper-V 役割インストール時のみ） |
| `S04_iis.json` | IIS 情報（Web Server 役割インストール時のみ） |
| `S05_dns.json` | DNS サーバー情報（DNS 役割インストール時のみ） |
| `S06_dhcp.json` | DHCP サーバー情報（DHCP 役割インストール時のみ） |
| `S07_fileserver.json` | ファイルサーバー情報（FS-FileServer 役割インストール時のみ） |
| `S08_rds.json` | RDS 情報（Remote Desktop Services 役割インストール時のみ） |

---

## スクリプトの使い方

```powershell
# 管理者権限で PowerShell を起動して実行
.\Collect-WindowsEnv.ps1

# 出力先を指定する場合
.\Collect-WindowsEnv.ps1 -OutputPath "D:\EnvBackup"

# HTML レポートも生成する場合
.\Collect-WindowsEnv.ps1 -GenerateHtml

# 出力先指定 + HTML レポート生成
.\Collect-WindowsEnv.ps1 -OutputPath "D:\EnvBackup" -GenerateHtml
```

`-GenerateHtml` を指定すると、収集した全ファイルの内容をまとめた `00_REPORT.html` が出力フォルダに追加生成されます。
左サイドバーにカテゴリナビゲーション、右側に各ファイルの内容が表示されます。

## 実行ポリシーエラーの回避方法

スクリプト実行時に以下のエラーが出る場合:
```
このシステムではスクリプトの実行が無効になっているため...
```

### 方法1: `-ExecutionPolicy Bypass` を付けて実行（推奨）

システムの設定を変更せず、このスクリプトの実行時のみポリシーを一時的に回避します。

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\Collect-WindowsEnv.ps1
```

### 方法2: 現在のセッションのみポリシーを変更

PowerShell ウィンドウを閉じると元に戻ります。

```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope Process
.\Collect-WindowsEnv.ps1
```

### 方法3: ユーザー単位でポリシーを変更（管理者権限不要）

```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
.\Collect-WindowsEnv.ps1
```

> **注意:** `Scope CurrentUser` の変更は現在のユーザーに永続的に適用されます。
> 不要になった場合は `Set-ExecutionPolicy -ExecutionPolicy Undefined -Scope CurrentUser` で元に戻してください。

出力先に `WindowsEnvAudit_YYYYMMDD_HHMMSS` フォルダが作成され、
カテゴリごとに JSON / テキストファイルが保存されます。

## 注意事項

### Windows Server での実行について

- **管理者権限必須**: サーバー役割の検出および S01〜S08 ステップの実行には管理者権限が必須
- **OS 自動判定**: OS の Caption に "Server" が含まれる場合を Windows Server として扱い、サーバー専用ステップを実行する
- **Get-WindowsFeature**: Windows Server 専用コマンド。クライアント OS では実行されない
- **ActiveDirectory モジュール**: `Get-ADDomain` 等の実行には ActiveDirectory モジュールが必要（AD DS インストールで自動追加、またはRSATで追加）
- **Hyper-V**: `Get-VM` 等は Hyper-V 役割インストール済みの場合のみ実行
- **IIS**: IISAdministration モジュール（IIS 10.0+、Windows Server 2016+）を優先し、なければ WebAdministration モジュールにフォールバック

- **管理者権限推奨**: BitLocker・ファイアウォールルール・secedit・デフォルトユーザープロファイルのレジストリハイブ読み取りには管理者権限が必要
- **ドメイン参加マシン**: GPO 情報・ドメインログオンスクリプト・フォルダリダイレクト等の組織管理情報が追加で取得される
- **セキュリティ情報**: MDE オンボーディング状態・Defender 設定・ASR ルール等はセキュリティ担当者向けの情報を含む
- **gpresult /user の制限**: 管理者権限で実行した場合に全ユーザー分の HTML レポートを生成するが、対象マシンに**一度もログインしていないドメインユーザー**はローカルプロファイルが存在しないため取得できない。`Win32_UserProfile` に登録されているローカルユーザーおよびログイン実績のあるドメインユーザーのみが対象となる
- **HKCU レジストリの他ユーザー分収集**: ログアウト済みユーザーの HKCU ポリシーは `NTUSER.DAT` を `reg load` でマウントして取得する。スキャン後は自動でアンロードするが、PowerShell がレジストリハンドルを保持している場合はアンロードに失敗することがある（スクリプト実行中に手動で対象ユーザーとしてログインが発生した場合など）
