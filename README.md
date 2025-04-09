# PDFcraft  
PDFcraft is a Python-based tool for manipulating PDF files — merging, splitting, watermarking, image insertion, and more.  
To prevent accidental file deletion, PDFcraft is designed to always generate new output files.  

PDFcraftは、PDFの結合・分割・透かし追加・画像挿入などを行うPython製ツールです。  
誤ってファイルを消すことを防ぐため、出力ファイルは常に新しく生成される仕様になっています。

---

## ✨ Features / 主な機能

- ✅ **Merge and split PDFs**  
　PDFの結合・分割

- ✅ **Add watermark text to each page**  
　各ページに透かし文字を追加

- ✅ **Insert JPG images into a PDF**  
　JPG画像をPDFに挿入

- ✅ **Extract or replace specific pages**  
　特定のページを抽出・差し替え

- ✅ **Support for scheduled and automated tasks**  
　スケジュール実行・自動化に対応

---

## 🚀 Getting Started / はじめかた

Place `PDFcraft.py` and `language.json` in the same folder and run:  
`PDFcraft.py` と `language.json` を同じフォルダに置いて、以下のように実行してください：

```bash
python PDFcraft.py
```

Or, if you're using the executable version:  
または、実行ファイル版を使う場合は：

```text
PDFcraft.exe と language.json を同じフォルダに置いてダブルクリックで実行してください。  
Place PDFcraft.exe and language.json in the same folder and double-click to run.
```

※ Windows 専用のGUIツールです。  
※ This is a GUI tool designed for Windows only.

---

## 🌐 Language Support / 言語対応

- English
- Japanese
- German
- French
- Spanish

The interface language is selected based on the `"language"` key in `language.json` (choose from `en`, `de`, `fr`, `es`, `ja`).  
If the file is missing or invalid, Japanese will be used as fallback.

表示言語は `language.json` の `"language"` キーによって選ばれます（`en`、`de`、`fr`、`es`、`ja` から選択）。  
ファイルが存在しない場合や読み込めない場合は、日本語が既定で使用されます。

---

## 📦 Requirements / 必要なパッケージ

For running from source:  
ソースから実行する場合の依存パッケージは以下の通りです。

See [`requirements.txt`](./requirements.txt) for installation.  
インストールには [`requirements.txt`](./requirements.txt) をご利用ください。

---

## 📜 License / ライセンス

This project is licensed under the  
[Creative Commons Attribution 4.0 International License (CC BY 4.0)](https://creativecommons.org/licenses/by/4.0/).  
本プロジェクトは  
[クリエイティブ・コモンズ 表示 4.0 国際ライセンス（CC BY 4.0）](https://creativecommons.org/licenses/by/4.0/) に基づき提供されています。

You may use, modify, and redistribute this tool, including for commercial purposes,  
as long as you give appropriate credit.  
商用利用・改変・再配布は自由ですが、著作者クレジットを明記してください。

---

## 👤 Author / 作者

**Kenji Niwa**  
[**koromokkuru lab.（コロモックル研究所）**](http://netyama.sakura.ne.jp/db/db.cgi?folder=kuruma)
---
