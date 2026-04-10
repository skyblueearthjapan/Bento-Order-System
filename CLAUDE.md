# お弁当予約アプリ (Bento Order System)

## 作業場所

- **GitHub リポジトリ:** https://github.com/skyblueearthjapan/Bento-Order-System.git
- **Google Apps Script:** https://script.google.com/u/0/home/projects/1GfiXEhnA2sW7Exok8z5Pur-S--2Zl6lCdD02uhT98yFFgLY-rjD6sox0/edit

## 関連ドキュメント

- **要件定義書:** `要件定義書.md`

## デプロイルール

- **デプロイ（Webアプリ公開）はユーザー（管理者）が実施する。Claudeはデプロイを実行しない**
- Claudeが行うのは `clasp push` のみ（コードをGASプロジェクトに反映するまで）
- デプロイURLの設定・共有範囲の変更もユーザーが実施

## コミット・プッシュルール（必須）

コード作成・修正後は、以下の手順を**必ず**実行すること：

1. **差分確認**: `git diff` でローカルの変更内容をしっかり確認する
2. **意図しない差分チェック**: 今回の修正以外にローカルデータとの意図していない差分が生じていないか確認する
3. **リポジトリ確認**: 正しいリポジトリ（Bento-Order-System）に対して操作しているか確認する
4. **GAS確認**: Google Apps Script側も同様に、正しいプロジェクトに対する変更か確認する
5. **コミット・プッシュ**: 上記すべて確認完了後、GitHubへコミット＆プッシュを行う
6. **clasp プッシュ**: GASコードは clasp push でデプロイする
