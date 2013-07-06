# importCsv
=========

ローカルのCSVファイルを、google spreadsheetの特定のシートにインポートします(google apps script)


## はじめに
google apps scriptの情報はそもそも少ないと思っていますが、その中でも<br>
ローカルのCSVファイルを読み込んで、シートに書き込む事例はかなり少なかったのでまとめてみました。<br>
（googleドライブ上のCSVファイルをインポートする例はまぁまぁあるんですけどね）<br>
<br>
各種事例の中でネックだったのは、SJISのCSVがNGだった事と、大きなサイズのCSVファイルだと遅すぎて<br>
実用性の面で難ありだったことです。<br>
ですので、これらを解決しました。<br>
<br>

## CSVファイルの文字コードについて

巷にあるgoogle apps scriptでCSVインポートする例は、CSVファイルがsjisだと文字化けするorフリーズしちゃいました。<br>
<br>
読み込んだテキストをsplitする前に、文字コード変換したり色々考えましたが、<br>
よくよくAPIのドキュメントとにらめっこすると、<br>
Class Blob - Google Apps Script - Google Developers : <br>
<https://developers.google.com/apps-script/reference/base/blob#getDataAsString()>
<br>
あれ、getDataAsString() charset引数受け付けてるやん。<br>
ということで、"Shift_JIS" 渡すだけでスンナリ通りました。(引数無しだとutf-8)<br>
<br>
## setValueによるセル書き込みが超遅いことについて
<br>
getRange() → setValue()で１セルずつ書き込むと、メチャクチャ遅いです。<br>
大きなCSVファイルだと、1ファイルをインポートするのに10分以上かかることもザラでした。<br>
でも、setValues()でまとめて書き込むと劇的に速くなります。<br>
<br>
しかし、setValues()には色々クセがあって結構扱いがメンドウなんです。<br>
あらかじめシートのサイズをデータの行列サイズよりも広げておかないと落ちるとか、<br>
書き込みセル数が251?以上だと落ちるとか。 詳しくは参考ページへ。<br>
<br>
ということで、そういった制限を考慮しながら１行ずつsetValues()する事で高速書き込み。<br>
<br>

## ソース抜粋
全ソースはgithub <https://github.com/prog-mitsu/importCsv> に上げてありますので、<br>
ご興味ある方は持って行って下さい。<br>
<br>

## 最後に

ローカルファイル選択アップロード<br>
SJIS対応<br>
高速書き込み<br>
が実現できたので、滅EXCEL、google spreadsheet推進への野望が一歩前進しました。<br>
<br>

## 参考ページ
Google Apps Script ってすごいね : 
<http://moblogger.r-stone.net/blogs/9016404448327222924/posts/1382923398397652155>

守破離でいこう!! : 
<http://ishikawa.r-stone.net/>

[GoogleAppsScript]setValuesではまったところ : minoawのブログ : <http://blog.livedoor.jp/minoaw/archives/1523932.html>

[GAS][スプレッドシート]処理速度を向上するには : 逆引きGoogle Apps Script : <http://www.bmoo.net/archives/2012/04/313959.html>


