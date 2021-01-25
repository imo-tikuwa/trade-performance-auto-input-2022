# trade-performance-auto-input-2021

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

## 自動化用の実行バッチ例
`trade-performance-auto-input-2021.bat`としてリポジトリ直下に置いてあります。  
(venv構築およびrequirements.txtに記載のモジュールのインストールを済ませておく必要あり)
```
@echo off

cd /D %~dp0
call .\venv\Scripts\activate.bat
python app.py --debug
call deactivate.bat
pause
```
