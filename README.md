# Windows 11 環境調査 - 調査項目リスト

## 目的
現在のWindows 11環境を一から再構築できるよう、設定・構成情報を網羅的に収集する。

---

## 調査カテゴリ一覧

### 1. システム基本情報
- OS バージョン・ビルド番号・エディション
- コンピュータ名・ワークグループ/ドメイン
- ハードウェア構成（CPU、RAM、GPU、マザーボード）
- BIOS/UEFI 情報（バージョン、セキュアブート設定）
- ディスク構成（パーティション、ファイルシステム）
- インストール日時

### 2. インストール済みソフトウェア
- Win32 アプリ（Add/Remove Programs 相当）
- MSI インストーラー直接インストール（ProductCode・UpgradeCode 付き詳細）
- ポータブルアプリ（インストーラーなし・EXE 直置き）のスキャン
  - スキャン対象: Program Files / Program Files (x86) / AppData\Local\Programs / デスクトップ / ダウンロード / ドキュメント / PortableApps
  - Win32 登録済みアプリとの照合結果（`IsRegistered` フラグ）を付与
- Microsoft Store アプリ (AppX/MSIX)
- winget パッケージ一覧
- Chocolatey パッケージ一覧（インストール済みの場合）
- Scoop パッケージ一覧（インストール済みの場合）

### 3. 開発ツール・ランタイム
- .NET ランタイム・SDK バージョン
- Node.js / npm バージョン・グローバルパッケージ
- Python バージョン・pip パッケージ
- Java / JDK バージョン
- Go / Rust / Ruby バージョン
- Git 設定（グローバル config）
- Docker / WSL2 設定
- WSL ディストリビューション一覧
- Visual Studio Code 拡張機能
- Visual Studio インストール済みコンポーネント

### 4. Windows 機能・オプション機能
- Windows オプション機能（Hyper-V、WSL、IIS 等）
- Windows Capabilities
- Windows Server Features（サーバーの場合）

### 5. サービス
- 実行中・停止中のサービス一覧
- スタートアップの種類（自動/手動/無効）
- 非標準サービス（Windows 標準以外）

### 6. ネットワーク設定
- ネットワークアダプター一覧・設定（IP/サブネット/ゲートウェイ/DNS）
- Wi-Fi プロファイル（SSID 一覧）
- hosts ファイル
- プロキシ設定
- ファイアウォールプロファイル設定
- ファイアウォールカスタムルール（インバウンド/アウトバウンド）

### 7. ユーザー・セキュリティ設定
- ローカルユーザーアカウント一覧
- ローカルグループ・メンバーシップ
- 環境変数（ユーザー/システム）
- UAC 設定レベル
- BitLocker 状態
- Windows Defender / セキュリティセンター設定
- 監査ポリシー

### 8. スタートアップ・スケジュールタスク
- スタートアップアプリ（レジストリ Run/RunOnce・スタートアップフォルダ）
- タスクスケジューラ登録タスク（カスタムタスク）

### 9. 電源・パフォーマンス設定
- 電源プラン（アクティブ・全プラン設定）
- スリープ・休止設定
- 仮想メモリ（ページファイル）設定
- 視覚効果設定

### 10. 地域・言語・入力設定
- タイムゾーン
- 地域・ロケール設定
- インストール済み言語パック
- キーボードレイアウト・入力メソッド（IME）

### 11. ディスプレイ・UI 設定
- 解像度・リフレッシュレート
- DPI スケーリング
- 夜間モード設定
- タスクバー設定
- スタートメニュー設定

### 12. ファイルシステム・共有設定
- 共有フォルダ一覧
- ドライブのマッピング（ネットワークドライブ）
- ジャンクション・シンボリックリンク
- NTFS 特殊アクセス権（重要フォルダ）

### 13. ドライバー
- インストール済みデバイスドライバー一覧
- 署名なし・問題のあるドライバー

### 14. 証明書
- 個人用・信頼済みルート証明書（カスタム追加分）

### 15. レジストリ（重要箇所）
- 自動起動キー（HKLM/HKCU Run）
- デフォルトプログラム・ファイル関連付け
- 環境変数（レジストリ）
- ポリシー設定キー

### 16. シェル・ターミナル設定
- デフォルトシェル設定
- Windows Terminal profiles.json
- PowerShell プロファイル内容
- PowerShell インストール済みモジュール
- エイリアス・関数（PowerShell プロファイル）

### 17. フォント
- インストール済みフォント一覧

### 18. プリンター・デバイス
- インストール済みプリンター
- デフォルトプリンター

### 19. ポリシー設定
- GPO 適用結果テキスト（`gpresult /r`）
- GPO 適用結果 HTML 詳細レポート（`gpresult /h`）
- レジストリポリシーキー値（`HKLM/HKCU:\SOFTWARE\Policies\`・`...\CurrentVersion\Policies\`）
- 監査ポリシー（`auditpol /get /category:*`）
- パスワード・アカウントロックアウトポリシー（`net accounts`）
- ローカルセキュリティポリシー全体（`secedit /export`、管理者権限時のみ）

### 20. スクリーンショット
- スタートメニューを開いた状態のデスクトップ
- 設定アプリを全画面で開いた状態

---

## スクリプトの使い方

```powershell
# 管理者権限で PowerShell を起動して実行
.\Collect-WindowsEnv.ps1

# 出力先を指定する場合
.\Collect-WindowsEnv.ps1 -OutputPath "D:\EnvBackup"
```

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
