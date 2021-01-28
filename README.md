# trade-performance-auto-input-2021

## このプログラムについて
こちらの成績管理ツール(Excelファイル)への入力を自動化します  
https://stock-analysis28.com/tool

## 環境構築
```
python -m venv venv
.\venv\Scripts\activate.bat
pip install -r requirements.txt
python app.py 
```

## オプション
| オプション名 | 内容 |
|---|---|
| --debug | デバッグログを出力します。 |

## 動作確認環境
 - Windows 10 Pro
 - python 3.8.2

## 使い方
1. venvの環境構築する
2. 15時以降にapp.pyを呼び出す
3. 初回のみchromedriverのパス、SBI証券のログインID/PW、追記先のExcelファイルパスを設定する
4. 2回目以降は自動化用のバッチ(`trade-performance-auto-input-2021.bat`)から起動するなど
