# プロデルSusieホストプラグイン(Produire.Susie)
画像ビューアであるSusieに対応したプラグインをプロデルから利用するためのプロデル用プラグインです。

◆動作環境情報◆
プロデル 1.9.1170

## 利用方法
Produire.Susie.dllを
プロデルのpluginsフォルダにコピーしてください。

もしくはプログラムと同じフォルダに置き、「利用する」文でプラグインを読み込みます。

```
「Produire.Susie.dll」を利用する
```

## 対応Susieプラグインについて
Susieに対応したプラグイン(.spi)が利用できます。

Susie API仕様は、次のアドレスに記載されている内容に準拠しています。
http://www2f.biglobe.ne.jp/~kana/spi_api/index.html

利用するには、プラグインのビット数とプロデルのビット数を対応させる必要があります。
プラグインが32bit向けの場合は、プロデルデザイナ(32ビット版)を利用してください。
実行可能ファイルは32ビット版(WoW64)として作成してください。

##　関連リンク(主なプラグインの配布場所)

「Susieの部屋」(公式サイト)
https://www.digitalpad.co.jp/~takechin/

「TORO's Library」
http://toro.d.dooo.jp/slplugin.html

これらのプラグインを利用したプログラムを配布する場合には
各プラグインのドキュメントのライセンスをご確認ください。


## 「Susieホスト」種類の手順と設定項目

次の手順と設定項目が利用できます。

### Susieホストへ【パス:文字列】を登録する

.spiファイルをホストに登録します。
ホストに登録されたプラグインからファイル形式に対応するプラグインが選択されます。

### Susieホストで【画像ファイル:文字列】を開く:画像

指定された画像ファイルを開き画像を取得します。

### Susieホストで【書庫ファイル:文字列】の一覧

指定された書庫ファイルを開き画像を取得します。

### Susieホストで【書庫ファイル:文字列】から【対象ファイル名:文字列】を【抽出先ファイル:文字列】へ抽出する

指定された書庫ファイルから対象ファイルを抽出し、抽出先フォルダへ保存します。

### Susieホストで【書庫ファイル:文字列】から【対象ファイル名:文字列】を抽出する:バイナリデータ

指定された書庫ファイルから対象ファイルを抽出し、バイナリデータとして返します。

### 「プラグイン一覧」設定項目

登録されたプラグインの一覧を返します。


## 「Susie書庫項目」種類の設定項目

### 「CRC」設定項目

CRCを返します。

### 「サイズ」設定項目

サイズを返します。

### 「パス」設定項目

書庫内での相対パスを返します。

### 「ファイル名」設定項目

ファイル名を返します。

### 「圧縮後サイズ」設定項目

圧縮後サイズを返します。

### 「圧縮方法」設定項目

圧縮方法を返します。

### 「位置」設定項目

位置を返します。

### 「更新日時」設定項目

位置を返します。


## 「Susieプラグイン」種類の設定項目

### [Susieプラグイン]が【ファイル:文字列】を対応するかどうか:真偽値

指定したファイルの扱いに対応しているかどうか

### [Susieプラグイン]にて〈【ウィンドウ部品】へ〉バージョン情報を表示する

プラグインのバージョン情報を表示します。
(プラグインが対応している場合のみ)

### [Susieプラグイン]にて〈【ウィンドウ部品】へ〉設定画面を表示する

プラグインの設定画面を表示します。
(プラグインが対応している場合のみ)

### 「APIバージョン」設定項目

APIバージョンを返します。

### 「対応ファイルパターン」設定項目

対応している形式の拡張子をパターンで返します。

### 「対応フィルタ」設定項目

開く選択画面の「フィルタ」設定項目の書式で対応する形式のフィルタを返します。

### 「名前」設定項目

プラグインの表示名を返します。


Copyright(C) 2022 utopiat.net. https://github.com/utopiat-ire/