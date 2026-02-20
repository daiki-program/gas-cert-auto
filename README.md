# gas-cert-auto
手動で行っていた修了証の作成・送付業務をGoogle Apps Script により自動化。自動送信後、管理者へ結果報告メールを送信。

### 📝 概要
Google Workspace（スプレッドシート、ドライブ、Gmail）を連携させた、社内向け研修修了証の自動発行・送付システムです。
手動で行っていた修了証の作成・送付業務をGoogle Apps Script (JavaScript) により自動化し、ヒューマンエラーの防止と工数削減を実現しました。

### 🚀 主な機能
* **自動PDF生成**: スプレッドシートの更新を検知し、ドキュメントテンプレートから個別の修了証PDFを生成。
* **バッチ送信**: 生成したPDFを翌日の指定時間に一括メール送信。
* **管理者報告機能**: 送信完了後、実行結果を管理者に自動レポート。

### 💻 使用技術
* **Language**: Google Apps Script (JavaScript)
* **Services**: Spreadsheet / Drive / Docs / Gmail
* **Triggers**: インストール型トリガー（Edit / Time-driven）

### 📄 設計書（ポートフォリオ詳細）
システムの詳細な仕様、業務フロー、シーケンス図については以下の設計書をご参照ください。
* [ポートフォリオ_DS.pdf](./ポートフォリオ_DS.pdf)


---
Developed by D. S. Spiderman
