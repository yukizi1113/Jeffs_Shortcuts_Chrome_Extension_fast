# Jeff's Shortcuts for Google Sheets - Fast

macOS版ChromeのGoogleスプレッドシートで、Windows版Excelに近いキーボード操作を使うためのChrome拡張機能です。

Chrome拡張がキー入力を直接受け取り、Google Sheets API v4を呼び出します。各スプレッドシートにApps Scriptを追加する必要はありません。付属のKarabiner-Elements設定を併用すると、Windows向けキーボードからCtrl中心の操作を使いやすくできます。

## 主な特徴

- `docs.google.com/spreadsheets/` で開くすべてのGoogleスプレッドシートに適用
- Apps ScriptやTool Finderを経由しない高速な処理
- 書式、色、罫線、行列サイズ、SuperFill、参照切替などをショートカットで実行
- 初回のみGoogleアカウントのOAuth承認が必要
- Y-U0036向けのWindows風Karabiner-Elements設定を同梱

## 必要なもの

- macOS
- Google Chrome
- Googleアカウント
- Googleスプレッドシートへのアクセス権
- Karabiner-Elements（付属のWindows風キーボード設定を使う場合）

## Chrome拡張機能のインストール

1. このリポジトリをダウンロードします。GitHubの **Code → Download ZIP** を選び、ZIPを展開してください。
2. Chromeで `chrome://extensions/` を開きます。
3. 右上の **デベロッパー モード** を有効にします。
4. **パッケージ化されていない拡張機能を読み込む** を選びます。
5. `manifest.json` が入っているリポジトリのフォルダを指定します。
6. Chromeのツールバーから **Jeff's Shortcuts** を開き、**Googleアカウントを承認** を押します。
7. Googleスプレッドシートを開くか再読み込みします。起動時の通知が表示されたら準備完了です。

更新版を取り込んだ場合は、`chrome://extensions/` の拡張機能カードにある再読み込みボタンを押し、開いているGoogleスプレッドシートも再読み込みしてください。

## OAuthについて

`manifest.json` にはGoogle Sheets / Drive用のOAuthクライアントIDが設定されています。承認に失敗する環境では、Google CloudでChrome拡張機能用のOAuthクライアントを作成し、`manifest.json` の `oauth2.client_id` を自分のクライアントIDへ置き換えてから拡張機能を再読み込みしてください。

要求する権限は次の用途で使用します。

- Google Sheets: セルの値・数式・書式・行列サイズなどの読み書き
- Google Drive: 日付付きコピーの作成

## Karabiner-Elements設定

設定ファイル: [`karabiner/Y-U0036-windows-like-shortcuts.js`](karabiner/Y-U0036-windows-like-shortcuts.js)

このファイルは、現在のKarabiner-Elements構成で使われている `eval_js` 形式のJavaScriptルールです。Y-U0036キーボードだけに適用される `device_if` 条件を含みます。

### 設定方法

1. Karabiner-EventViewerを開き、対象キーボードの `vendor_id` と `product_id` を確認します。
2. 設定ファイル冒頭の `VENDOR_ID` と `PRODUCT_ID` が実機の値と異なる場合は、自分の環境用コピーで値を変更します。
3. Karabiner-ElementsのComplex ModificationsでJavaScriptルールを追加し、ファイル全体を `eval_js` の内容として登録します。
4. `Y-U0036 - Windows-like shortcuts` ルールを有効にします。
5. ChromeでGoogleスプレッドシートを使うときは、`Ctrl + Option + Shift + 0` で **Jeff's Sheets mode** をONにします。同じキーでもう一度押すとOFFになります。

> 設定ファイル内の説明コメントにはF12と書かれている箇所がありますが、現在の実装上の切替キーは `Ctrl + Option + Shift + 0` です。

Jeff's Sheets modeがONの間は、対象キーボードからのJeff用ショートカットをControl / Option / ShiftのままChromeへ渡します。OFFの間は、通常のWindows風Ctrl→Command変換が優先されます。Jeffモードの対象ブラウザはChromeです。

### Karabiner設定に含まれる主なWindows風操作

- `Alt + Tab`: アプリ切替
- `Ctrl + C/V/X/Z/A/S/F/P` など: macOSのCommandショートカットへ変換
- Finderの `Ctrl + X` → `Ctrl + V`: ファイル移動
- Finderの `Enter`: 開く
- Finderの `F2`: 名前変更
- Finderの `Delete`: ゴミ箱
- `Windows + L`: 画面ロック
- `Windows + E`: Finder
- `Windows + R`: Spotlight
- `Windows + Shift + S`: 範囲スクリーンショットをクリップボードへ保存
- `Print Screen`: 画面全体をクリップボードへ保存
- Chrome / Edgeの `F5`: 再読み込み
- Chrome / Edgeの `F12`: DevTools
- Jeffモード中の `Ctrl + 矢印`: Googleスプレッドシートのデータ領域端へ移動

## ショートカット一覧

`Ctrl` はControl、`Alt` はOption、`Cmd` はCommandを表します。

| ショートカット | 動作 |
|---|---|
| `Ctrl + Shift + 1` | Excel風の桁区切り・小数2桁 |
| `Ctrl + Shift + 5` | Excel風のパーセント・小数0桁 |
| `Ctrl + [` | 参照元セルへ移動 |
| `Ctrl + Shift + N` | 数値書式を切替 |
| `Ctrl + Shift + C` | 文字色を切替 |
| `Ctrl + Shift + V` / `Cmd + Shift + V` | 塗りつぶし色を切替 |
| `Ctrl + Shift + S` | 日付付きDriveコピー |
| `Ctrl + Shift + D` | SuperFillを下方向へ実行 |
| `Ctrl + Alt + D` | 空白までSuperFillを下方向へ実行 |
| `Ctrl + Shift + Z` | SuperFillを元に戻す |
| `Ctrl + Shift + R` | SuperFillを右方向へ実行 |
| `Ctrl + Alt + R` | 空白までSuperFillを右方向へ実行 |
| `Ctrl + Shift + Q` | 選択範囲の数式を再評価 |
| `Ctrl + Shift + F4` | 参照を絶対参照へ変換 |
| `Ctrl + Alt + F4` | 参照を相対参照へ変換 |
| `Ctrl + Shift + E` | 中央揃え |
| `Ctrl + Alt + Shift + E` | 中央揃え |
| `Ctrl + Alt + Shift + C` | 文字色をクリア |
| `Ctrl + Alt + Shift + V` | 塗りつぶしをクリア |
| `Ctrl + Shift + W` | 折り返しを切替 |
| `Ctrl + Shift + M` | インデント（現在は変更なし） |
| `Ctrl + Alt + Shift + M` | インデント（現在は変更なし） |
| `Ctrl + Shift + T` | TableTop風書式を切替 |
| `Ctrl + Shift + I` | 指定ネイビー書式を切替 |
| `Ctrl + Alt + Shift + F` | Arial 10・グリッド線非表示 |
| `Ctrl + Shift + B` | 罫線を切替 |
| `Ctrl + Alt + Shift + B` | 罫線をクリア |
| `Ctrl + Alt + Shift + W` | 列幅を既定値へ設定 |
| `Ctrl + Alt + Shift + H` | 行高を既定値へ設定 |
| `Ctrl + Alt + Shift + → / ←` | 列幅を増減 |
| `Ctrl + Alt + Shift + ↓ / ↑` | 行高を増減 |
| `Ctrl + Alt + W` | 列幅を正方形用サイズへ設定 |
| `Ctrl + Alt + L` | 垂直方向の配置を切替 |
| `Ctrl + Alt + H` | 水平方向の配置を切替 |
| `Ctrl + Shift + , / .` | 小数桁を減らす / 増やす |
| `Ctrl + Alt + Shift + O` | 診断メッセージを表示 |
| `Ctrl + Shift + K / J / H` | Chromeタブを拡大 / 縮小 / 100% |
| `Ctrl + Alt + 1` | F1ヘルプのブロックを切替 |

## ファイル構成

| ファイル | 内容 |
|---|---|
| `manifest.json` | Chrome拡張機能の権限・OAuth・読み込み設定 |
| `background.js` | Google Sheets / Drive APIを呼び出すバックグラウンド処理 |
| `content.js` | Googleスプレッドシート上のキー入力検出 |
| `popup.html`, `popup.js` | 初回OAuth承認用のポップアップ |
| `content_old.js` | 旧版の保存用ファイル。現在のmanifestからは読み込まれません |
| `README_JA.txt` | 元の簡易説明 |
| `karabiner/Y-U0036-windows-like-shortcuts.js` | Y-U0036向けWindows風Karabiner設定 |

## 注意事項

- 選択範囲の取得にはGoogleスプレッドシート画面の `.waffle-name-box` を使用しています。Google側の画面構造が変わると、将来修正が必要になる可能性があります。
- セルの読み書き自体はGoogle Sheets API v4を使用します。
- Google Sheets APIには非破壊のセルインデントAPIがないため、現在インデント用ショートカットは何も変更しません。
- 書式やDriveコピーを実行できる権限のあるGoogleアカウントで承認してください。
- このリポジトリは現状の設定を共有するためのものです。利用前に内容と権限を確認してください。
