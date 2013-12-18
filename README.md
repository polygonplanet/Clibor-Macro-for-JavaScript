Clibor-Macro-for-JavaScript
===========================

## これはなに？

クリップボード履歴ソフト [Clibor][1] のマクロを JavaScript で書けるようにするパッチです。

## 必要なもの

 1. Clibor 本体
 2. Python (公式のマクロ)
 3. pywin32 (公式のマクロ)


一部の機能で Python を使用しています。  
Python を入れてない場合、[公式の解説][2] 通りにインストールしてください。

## ダウンロード

[Clibor-Macro-for-JavaScript.zip][5]


## インストール

 1. 「clb_ex.bat」が Clibor.exe と同階層にある場合、適当な名前に変えるなどしてバックアップします。  

 2. ダウンロードした本パッチの「clibor-patch」フォルダ内の全ファイル
    * clb_ex.bat
    * jsmacro.clip.bat
    * jsmacro.clip.py
    * jsmacro.trigger.wsf

 を、Clibor.exe と同階層にコピーします。
 3. 「jsmacro.clip.bat」をエディタ等で開き、

```dos
rem ---Set your python directory---
set PYTHON_DIR=C:\Python27
rem -------------------------------
```

 3行目あたり、上の Python パスを自分の環境に合わせます。


以上でセットアップは完了です。

実際に動くかテストしてみましょう。

タスクトレイ Clibor を右クリックしメニューを開き、
「定型文の編集」でマクロ作成ウィンドウを開きます。

「新規作成」で編集画面を開き、以下のコードを貼り付けて保存します。

```javascript
#<$C_CLB_PYTHON/>
alert(' hello   world  !'.split(/\s+/).filter(function(s) {
  return s;
}).map(function(s) {
  return s.charAt(0).toUpperCase() + s.slice(1);
}).join(' '));
```

1行目は、本来 Python で動くのをトラップするために必要です。

`'Hello World !'` とアラートが表示されれば成功です。

※環境により、alert が最前面にならないかもしれません

## JavaScript 環境について

WSH, JScript で実行しています。

JavaScript としては非常に古く、ECMAScript の仕様にもあまり準じていないでしょう。

本パッチでは、いくつかの polyfill により、ECMAScript 5th ぽいことは、あらかじめ使用できるようになっています。(Array.prototype.forEach 等)

マクロのスコープ内は、以下のオブジェクト/関数が定義されています。


 - print : function (text) : textをアクティブウィンドウに貼り付けます
 - sleep : function (msec) : msec(ミリ秒)待機します
 - Clipboard : object { : クリップボードを扱うオブジェクト

  Clipboard:

    - get : function () : 現在のクリップボードの内容を返します
    - set : function (text) : クリップボードにtextを設定します
    - copy : function () : アクティブウィンドウに対しコピー(Ctrl+C)キーを送ります
    - cut : function () : アクティブウィンドウに対しカット(Ctrl+X)キーを送ります
    - paste : function () : アクティブウィンドウに対しペースト(Ctrl+V)キーを送ります

 }
 - withClipboard : function (fn) : 関数fnをClipboardのもとで実行します(this=Clipboard)
 - getSelectedText : function () : アクティブウィンドウで選択されているテキストを返します
 - LocalFile : constructor function (fileName) { : ローカルファイルを扱うコンストラクタ

 LocalFile.prototype:

    - create : function (overwrite = true) : 新しいファイルを作ります
    - exists : function () : ファイルが存在すれば true を返します
    - remove : function (force = true) : ファイルを削除します
    - getSize : function () : ファイルサイズ(byte)を返します
    - copy : function (to, overwrite = true) : ファイルをtoへコピーします
    - move : function (to) : ファイルをtoへ移動します
    - read : function (charset = 'UTF-8') : ファイルの内容をすべて読み込んで返します
    - write : function (text, charset = 'UTF-8') : ファイルにtextを書き込みます

 }
 - gc : function () : 強制的に gc (garbage collect) します。通常は必要ありません


以下は標準のDOM と同じに扱えます ( IE仕様 ※1 )

 - alert : function (msg)
 - confirm : function (msg)
 - setTimeout : function (func, delay)
 - clearTimeout : function (id)
 - setInterval : function (func, interval)
 - clearInterval : function (id)
 - XMLHttpRequest : constructor function ()


※1: IE仕様なのは、ActiveXObject('htmlfile') を使っているためです


## サンプル

CSSのようなハイフン(-)区切りをキャメルケースにする

```javascript
#<$C_CLB_PYTHON/>
// camelize
print(getSelectedText().toLowerCase().replace(/-(.)/g, function(a, b) {
  return b.toUpperCase();
}));
```

例:

`list-style-type` → `listStyleType`

----
選択範囲のすべての空白を切り詰め

```javascript
#<$C_CLB_PYTHON/>
// 選択範囲のすべての空白を切り詰め
withClipboard(function() {
  sleep(200);
  this.copy();
  var text = this.get();
  text = text.replace(/[ \t\u3000]{2,}/g, ' ').replace(/(?:\r\n|\r|\n){2,}/g, '\n');
  print(text);
});
```

例:

    hoge       fuga            ho
                 hgeo  ghe  geg
     gegg       foge


    piyo            hoge

↓

    hoge fuga ho
    hgeo ghe geg
    gegg foge
    piyo hoge

----
README (今見てるこれ) を表示する

```javascript
#<$C_CLB_PYTHON/>
// Open README
var url = 'https://github.com/polygonplanet/Clibor-Macro-for-JavaScript/blob/master/README.md';
var req = new XMLHttpRequest();

req.open('GET', url, false);
req.send(null);

var content = req.responseText;
var readmeHTML = content.match(/<article[^>]*>([\s\S]*?)<\/article>/i)[1];
var readmeText = readmeHTML.replace(/<\S[^>]*>/g, '');
alert(readmeText);
```

----
HTMLタグの除去 (少し細かく)

```javascript
#<$C_CLB_PYTHON/>
var text = getSelectedText();

text = [
  [/<!-*[\s\S]*?->|<!\s*\w+[^>]*>/g, ''],
  [/<\s*(\w+)\b[^>]*>([\s\S]*?)<\s*\/\s*\1\s*>/g, ' $2 '],
  [/<[^>]*>|<[![\]-]*|[-[\]]*>/g, '']
].reduce(function(html, item) {
  return html.replace(item[0], item[1]);
}, text);

print(text);
```

例:

```html
<html><head>
  <meta charset="utf-8">
  <title>タイトル</title>
</head>
<body>
  <div>こんにちは！</div>
  <script language="日本語">ようこそ！私のホームページへ！</script>
</body>
</html>
```

↓

    タイトル
    
    こんにちは！
    ようこそ！私のホームページへ！


----
選択範囲のコードをJSLint

```javascript
#<$C_CLB_PYTHON/>
/*JSLINT*/

var source = getSelectedText();
if (!source.trim()) {
  return;
}

var JSLINT_URL = 'http://jslint-for-wsh.googlecode.com/svn/trunk/jslint-for-wsh.js';
var JSLINT_FILE = JSLINT_URL.split('/').pop();

var options = {
   windows    : true,  // WScript is allowed
   white      : true,  // true: "sloppy" whitespace is ok
   plusplus   : true,  // true == ok to use ++
   properties : false, // do not barf on any undeclared properties
   passfail   : false, // do not stop after first error
   radix      : true   // do not puke on parseInt() with no radix
};

var code;
var file = new LocalFile(JSLINT_FILE);
if (!file.exists()) {
  // ファイルがなければダウンロードして保存
  file.write(code = download(JSLINT_URL));
} else {
  code = file.read();
}
file = null;

eval(code);
code = null;

alert(lint(source));


function download(url) {
  var req = new XMLHttpRequest();
  req.open('GET', url, false);
  req.send(null);
  return req.responseText.replace(
    /([{}()\s]+;)\s*\(\s*function\s*[()\s]+\{[\s\S]*[}();\s]+$/,
    '$1'
  );
}

function lint(code) {
  if (!JSLINT(code, options)) {
    return JSLINT.errors.reduce(function(err, e) {
      return (e && err.push([
        'line: ' + (e.line || 0),
        'pos: ' + e.character,
        'JSLINT: ' + e.reason +
        '\n    ' + (e.evidence || '').trim()
      ].join(', '))), err;
    }, []).join('\n');
  }
  return 'no error';
}
```

例:

下のコードを選択して実行

```javascript
var a = 1, 2, 3;
```

↓


```javascript
line: 1, pos: 12, JSLINT: Expected an identifier and instead saw '2'.
    var a = 1, 2, 3;
line: 1, pos: 12, JSLINT: Stopping.  (100% scanned).
```



## License

#### MIT

## Authors

 - [polygon planet][3] (twitter: [polygon_planet][4])

  [1]: http://www.amunsnet.com/soft.html
  [2]: http://www.amunsnet.com/soft/clibor_macro.html
  [3]: http://polygonpla.net/
  [4]: http://twitter.com/polygon_planet
  [5]: https://github.com/polygonplanet/Clibor-Macro-for-JavaScript/archive/master.zip

