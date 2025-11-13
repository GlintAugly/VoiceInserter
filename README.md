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

* Windows版
0. https://github.com/GlintAugly/VoiceInserter/releases からVoiceInserter_install.zipをダウンロード、展開してReleaseフォルダを開く
1. init.batを**管理者として**実行する
2. ダウンロード途中、Voicevoxの利用規約が出てくるので、内容をよく読み、同意出来る場合は同意する。(同意しない場合でも、既存音声ファイルの配置機能は利用可能です)

# 利用方法

0. DaVinci Resolveでプロジェクトを新規作成or開く
1. 上部メニュー→[ワークスペース]→[スクリプト]→[VoiceInserter]を選択
2. キャラ名を入力し「作成」を押す
3. ウィンドウが出てくるので、各種パラメーターを変更して「挿入」ボタンをクリック

# Lisence

This project is licensed under the MIT License, see the LICENSE.txt file for details

## Release同梱物のLisenceについて

### download-windows-x64.exe

Copyright (c) 2021 Hiroshiba Kazuyuki

Permission is hereby granted, free of charge, to any person obtaining
a copy of this software and associated documentation files (the
"Software"), to deal in the Software without restriction, including
without limitation the rights to use, copy, modify, merge, publish,
distribute, sublicense, and/or sell copies of the Software, and to
permit persons to whom the Software is furnished to do so, subject to
the following conditions:

The above copyright notice and this permission notice shall be
included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.