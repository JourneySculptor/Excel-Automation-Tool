# Excel-Automation-Tool (エクセル自動化ツール)

## 概要
**Excel-Automation-Tool** は、Python を使用して Excel データ処理を自動化するツールです。データのフィルタリング、レポートの生成、Googleスプレッドシートとの連携を通じて、業務効率を向上させることを目的としています。

**このプロジェクトのドキュメントは多言語対応しています**
- [English README.md](README.md) - English documentation.
- [日本語 README_ja.md](README_ja.md) - 日本語ドキュメント。

## 特徴
- **ユーザー入力対応**: ユーザーが指定したExcelファイルを処理可能
- **データフィルタリング**: 売上金額の閾値を基準にデータを自動フィルタリング
- **グラフ生成**: 売上データをカテゴリ別に可視化する棒グラフを作成
- **PDFレポート出力**: 売上データをPDFレポートとして保存
- **Googleスプレッドシート連携**: フィルタリングしたデータをGoogleスプレッドシートにアップロード

## プロジェクト構成
```plaintext
Excel-Automation-Tool/
│── sample_data/         # サンプルの Excel ファイルと出力ファイル
│   ├── sample.xlsx      # ユーザーが提供するサンプル Excel ファイル
│   ├── high_sales.xlsx  # フィルタリング後のデータ
│   ├── sales_chart.png  # 生成された売上グラフ
│   ├── sales_report.pdf # 生成された PDF レポート
│── venv/                # 仮想環境（GitHub には含まれない）
│── main.py              # メインスクリプト
│── requirements.txt     # 依存関係リスト
│── README.md            # 英語版ドキュメント
│── README_ja.md         # 日本語版ドキュメント
│── credentials.json     # Google Sheets API の認証情報（GitHub には含めない）
```

## 必要環境
このツールを使用するには、以下のライブラリが必要です。
```bash
pip install pandas openpyxl matplotlib fpdf gspread oauth2client
```

## インストールと使用方法
1. リポジトリをクローンする:
```bash
git clone https://github.com/JourneySculptor/Excel-Automation-Tool.git
cd Excel-Automation-Tool
```
2. 依存関係をインストール:
```bash
pip install -r requirements.txt
```
3. Excelファイルを `sample_data/` フォルダに配置する。
4. スクリプトを実行:
```bash
python main.py
```
5. 指示に従ってExcelファイルを処理。

## 出力例
### フィルタリング後のデータ:
売上が10,000円以上のデータを抽出し、`high_sales.xlsx` に保存。

### 生成されたグラフ:
カテゴリ別の売上を可視化した棒グラフが `sales_chart.png` に保存。

### PDFレポート:
売上データのサマリーレポートを `sales_report.pdf` に出力。

### Googleスプレッドシートへのアップロード:
正しく設定されている場合、フィルタリングされたデータがGoogleスプレッドシートにアップロードされる。

## Googleスプレッドシートの設定

Googleスプレッドシートとの連携を有効にするには、以下の手順を実施してください。

1. Google Cloudプロジェクトを作成し、Google Sheets APIを有効化。
2. サービスアカウントキー (credentials.json) を生成。
3. `credentials.json` をプロジェクトのルートディレクトリ (`Excel-Automation-Tool/`) に配置。
4. `main.py` を実行すると、自動的にデータがGoogleスプレッドシートにアップロード。

> **注意:`credentials.json` をGitHubにアップロードしないでください！`gitignore` に追加し、セキュリティを確保してください。** 

---

## 今後の改善予定
- 非技術者向けのGUIの実装
- レポートを自動送信するメール機能の追加
- エラーハンドリングを強化し、ユーザビリティ向上

> 💡 開発への貢献歓迎！
機能追加のアイデアやバグ報告は [GitHub の Issues](https://github.com/JourneySculptor/Excel-Automation-Tool/issues) で受け付けています。

## ライセンス
このプロジェクトはMITライセンスのもとで提供されます。
