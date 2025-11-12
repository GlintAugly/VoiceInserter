# VoiceInserter

DaVinci Resolve向けにvoicevoxなどの音声を画像・字幕と一緒に配置するスクリプト。
GUIで画像や字幕のプロパティを設定することで、共通した設定を引き継ぎながら、随時的に音声を組み込むことができます。

# 前提条件

* DaVinci Resolveのインストール
https://www.blackmagicdesign.com/jp/products/davinciresolve/　から、無料でインストール可能です。

* Pythonのインストール
https://www.python.org/downloads/ から、無料でインストール可能です。

※DaVinci Resolveでpythonを使用するための事前準備に関しては、後段のインストールバッチで自動的に行われます。また、事前に行っていた場合でも、支障はないはずです。

# できること

いわゆるボイロ実況などを作る際に必要になる、音声・字幕・キャラ画像を、音声の長さに合わせて配置します。
* 既存の音声ファイルと文字列、画像を、DaVinci Resolveで現在表示しているタイムライン上に配置する。
* Voicevoxを利用して音声を生成し、生成した音声と字幕を画像と一緒にタイムライン上に配置する。
* キャラごとに画像位置、文字位置、文字色などを設定し、反映する。

# インストール方法

0. voicevox_coreのダウンローダー(download-windows-x64.exe)を、https://github.com/VOICEVOX/voicevox_core/releasesからダウンロードしてくる。
0. init.batと同じフォルダにVoiceInserter.pyとvoicevox_coreのダウンローダーを配置する。
1. init.batを**管理者権限で**実行する。
2. ダウンロード途中、Voicevoxの利用規約が出てくるので、内容をよく読み、同意出来る場合は同意する。(同意しない場合でも、既存音声ファイルの配置機能は利用可能です)

# 利用方法

0. DaVinci Resolveでプロジェクトを新規作成or開く
1. 上部メニュー→[ワークスペース]→[スクリプト]→[VoiceInserter]を選択
2. キャラ名を入力し「作成」を押す
3. ウィンドウが出てくるので、各種パラメーターを変更して「挿入」ボタンをクリック

# Lisence

This project is licensed under the MIT License, see the LICENSE.txt file for details
