# make_PowerPoint

Wordファイルから、PowerPointを自動生成するWebアプリです。

PowerPoint資料作成の手間を軽減することを目的として開発しました。

【デモURL】
https://make-powerpoint.onrender.com/

---

## アプリ概要

Wordファイルをアップロードすると、内容を解析してPowerPointファイルを自動生成します。

作成したPowerPointはブラウザからダウンロードできます。

---

## 開発背景

業務で資料作成を行う際、

- Wordで原稿を作成
- PowerPointへ転記
- レイアウト調整

という作業に時間がかかることがありました。

そこで、Wordの内容からPowerPointを自動生成するツールを開発しました。

単なるスクリプトではなく、誰でも利用できるようにFlaskでWebアプリ化しています。

---

## 主な機能

### Wordファイルのアップロード

- .docxファイルに対応

### PowerPoint自動生成

- 1行目をタイトルスライドとして配置
- 2行目以降を本文としてスライド生成
- 空白行でスライドを区切る

### ダウンロード機能

- 生成したPowerPointをダウンロード可能
- ファイル名の指定に対応

### エラーハンドリング

- 不正ファイルのアップロードを検知
- エラーメッセージを表示

---

## 使用技術

| 分類 | 技術 |
|--------|--------|
| 言語 | Python 3 |
| Webフレームワーク | Flask |
| Word操作 | python-docx |
| PowerPoint操作 | python-pptx |
| WSGIサーバー | Gunicorn |
| デプロイ | Render |
| バージョン管理 | Git / GitHub |

---

## システム構成

```text
ユーザー
    ↓
Flask
    ↓
Wordファイル解析
    ↓
PowerPoint生成
    ↓
ダウンロード