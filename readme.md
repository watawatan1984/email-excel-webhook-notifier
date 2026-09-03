# メール添付Excel自動処理・REST API/Teams/WebPush通知システム 保守・運用マニュアル

このドキュメントは、本システムの概要、導入方法、日常の運用、およびトラブルシューティングについて解説する管理者向けマニュアルです。
Microsoft Graph API移行に対応しています。

## 1. システム概要
本システムは、メールに添付されたExcelシートから情報を取得・解析し、REST API経由でのデータ登録、Microsoft Teamsへの処理結果通知、およびWebPush（OneSignal）通知を自動実行するPythonスクリプトです。

### 処理フロー
1.  **メール受信**: サーバー上で定期実行され、指定された差出人からのメールをMicrosoft Graph API経由で取得します。
2.  **Excel抽出**: メールに添付されたExcelファイル（パスワード付き対応）をダウンロードし、復号・解析します。
3.  **データ連携**:
    *   **REST API (WordPress等)**: 抽出したデータをREST API経由で登録します。
    *   **Microsoft Teams**: 処理結果のサマリーをTeamsチャネルに通知します。
    *   **WebPush (OneSignal)**: スマートフォンアプリ向けにプッシュ通知を送信します。
4.  **メール既読化**: 処理が完了したメールを「既読」にし、次回実行時に重複処理しないようにします。

---

## 2. 動作環境・必須要件
*   **OS**: Linux (レンタルサーバー / クラウド環境)
*   **言語**: Python 3.x
*   **必須ライブラリ**:
    *   `msal` (Microsoft Authentication Library)
    *   `requests`
    *   `openpyxl`
    *   `msoffcrypto-tool`

---

## 3. セットアップ手順（初回・移行時）

### 3.1 必要なライブラリのインストール
サーバーにSSH等で接続し、以下のコマンドを実行してください。
```bash
pip3 install msal requests openpyxl msoffcrypto-tool
```

### 3.2 設定ファイル (`config.ini`) の準備
`config.ini` ファイルに環境に応じた設定を記述します。Azure ADの情報はAzureポータルから取得してください。

```ini
[Email]
# Microsoft Graph API連携用情報（Azure AD）
TENANT_ID = <AzureのテナントID>
CLIENT_ID = <AzureのクライアントID>
CLIENT_SECRET = <Azureのクライアントシークレット>
TARGET_EMAIL_ADDRESS = <受信対象のメールアドレス (例: info@example.com)>
TARGET_SENDER = <監視する差出人メールアドレス>

[Excel]
PASSWORD = <添付Excelのパスワード>
# ... (その他の設定は環境に合わせて設定)
```

---

## 4. 日常の運用・実行方法

### 4.1 手動実行
動作確認や緊急時に手動で実行する場合です。
```bash
# プログラムのディレクトリへ移動
cd /path/to/automation_directory

# 実行
python3 main.py
```
実行後、画面（またはログ）に `スクリプトを終了します。` と表示されれば完了です。

### 4.2 自動実行（Cron）
通常はCronにより定期実行します。設定状況の確認コマンド:
```bash
crontab -l
```
※ Cronの設定行の例: `0,30 * * * * /usr/bin/python3 /path/to/automation_directory/main.py`

---

## 5. 保守・トラブルシューティング

### 5.1 ログの確認
動作がおかしい、データが登録されない等は、まずログファイルを確認してください。
*   **ログ保存場所**: `logs/app.log`
*   **確認コマンド例**: `tail -f logs/app.log` (リアルタイム監視)

### 5.2 よくあるエラーと対処

#### Q1. 「Authorization_IdentityNotFound」等の認証エラーが出る
*   **原因**: Azure ADの設定 (`TENANT_ID`, `CLIENT_ID`) が間違っているか、ターゲットメールアドレスが存在しません。
*   **対処**: `config.ini` の内容とAzureポータルの登録情報が一致しているか確認してください。

#### Q2. 「InvalidCientSecret」エラーが出る
*   **原因**: クライアントシークレット（パスワード）の有効期限が切れています。
*   **対処**:
    1.  Azureポータルで新しいクライアントシークレットを発行してください。
    2.  `config.ini` の `CLIENT_SECRET` を新しい値に書き換えてください。

#### Q3. Excel添付なしのメールで止まる
*   **症状**: 添付がない場合でも「添付なし」としてログに残し、**自動的に既読（処理済み）** にします。
*   **確認**: `logs/app.log` に `Excel添付なし ... 処理済みとしてマークします` と記録されているか確認してください。

### 5.3 メンテナンス時の注意点
*   **パスワード変更**: Excelのパスワードが変わった場合は、`config.ini` の `[Excel] PASSWORD` を変更してください。
*   **ディレクトリ構造**:
    *   `main.py`: メインプログラム
    *   `config.ini`: 設定ファイル
    *   `processed_uids.txt`: 処理済みのメールIDを記録するファイル（テキストエディタで開いて行を消すと、過去のメールを再処理できます）。
    *   `logs/`: ログフォルダ。容量がいっぱいにならないよう、定期的に古いログを削除することをお勧めします。

---

## 6. お問い合わせ・サポート
本システムのコード修正や設定変更が必要な場合は、システム管理者へ連絡してください。