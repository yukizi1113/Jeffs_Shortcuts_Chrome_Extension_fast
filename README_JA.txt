Jeff's Shortcuts for Google Sheets - Fast
=========================================

目的
----
Apps Script / Tool Finderを介さず、Chrome拡張がGoogle Sheets上のキー入力を直接捕捉し、
Sheets API v4を直接呼ぶ高速版です。

・一度Chromeにインストールすれば、docs.google.com/spreadsheets/ で開く全Spreadsheetに適用
・各SpreadsheetへのApps Script貼付けは不要
・Hammerspoonも原則不要（Chromeに届かないキーだけ、必要なら後で補助に使える）
・初回だけGoogle OAuth承認が必要

重要
----
この拡張はGoogle Sheetsの「Name box」のDOMクラス .waffle-name-box から現在の選択範囲を取得します。
これはGoogle Sheetsの公開APIではなくWeb UIの実装詳細なので、GoogleがUIを変更すると将来修正が必要になる可能性があります。
一方、セルの読み書き自体は正式なGoogle Sheets API v4を使います。

OAuth client_id
---------------
manifest.json の
REPLACE_WITH_YOUR_CHROME_EXTENSION_OAUTH_CLIENT_ID.apps.googleusercontent.com
を、Google Cloudで作成した「Chrome extension」OAuthクライアントIDに置き換えてください。

現在の移植範囲
--------------
11 Number format toggle
12 Font color toggle
13 Fill color toggle
17 Dated Drive copy
18/19 SuperFill down
21 SuperFill undo
24/25 SuperFill right
26 formula re-trigger
27/28 references absolute/relative
31/32 center approximation
33/34 clear font/fill
35 wrap
36/37 indent: API不存在のためno-op
38 TableTop approximation
39 specified navy style
41 Arial 10 + gridlines off
42/43 borders
44-49 dimensions
51-55 square width / alignment / decimals
59 diagnostics
Zoom in/out/100%: Chrome tab zoom
F1 block toggle

