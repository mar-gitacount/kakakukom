# README

## 概要

このプロジェクトは、価格コムからデータを抽出し、そのデータをエクセルファイルに保存する `webdriver2.py` と、そのスクリプトを実行する`sub.py` と`subtest.py` から構成されています。

## ファイル構成

- `sub.py`:価格コムのサイトページ数に合わせて、webdriver2.pyを実行する。
- `subtest.py`:ページ取得をページ数単体で入力し、webdriver2.pyを実行する。sub.pyでエラーになった場合、にそのエラーになった番号を入力し、それを実行する。
- `webdriver2.py`: 価格コムからデータをスクレイピングし、エクセルに抽出する。

## 必要なライブラリ

以下のPythonライブラリが必要です。これらは `requirements.txt` ファイルに記載されています。

- requests==2.26.0
- beautifulsoup4==4.10.0
- pandas==1.3.4
- openpyxl==3.0.9

インストール方法:
```sh
pip install -r requirements.txt
