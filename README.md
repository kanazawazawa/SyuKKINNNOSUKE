# SyuKKINNNOSUKE
PC起動時に勤之助の出社ボタンを自動で押すアプリケーションです。  
C# Selenium Chrome
## Install
発行して生成される SyuKKINNNOSUKE.exe と chromedriver.exe を任意のフォルダに置いておきます。  
ファイル名を指定して実行 shell:startup で開かれるフォルダに SyuKKINNNOSUKE.exe のショートカットを置いておきます。  
.Net Core 3.0 のプレビュー機能 Single-file Publish を使っているので、必要なのは2つのexeのみです。
## 使い方等
初回のみ、お客様ID、ログインID、パスワードを入力します。ユーザー環境変数に保存しているので次回以降入力不要です。  
コマンドライン引数に clear を渡すとユーザー環境変数をクリアします。
## 動作確認環境
下記環境で動作確認しています。  
Windows 10 Pro
## 開発環境
Visual Studio 2019  
.Net Core 3.0
