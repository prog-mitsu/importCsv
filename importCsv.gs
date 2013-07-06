/**
 * @google spreadsheet用 CSVファイルインポート<br>
 * @author ms32
 * @version 1.0.0
 * 
 * setValueで１セルずつ書き込むと超遅いので、setValuesで書き込みます<br>
 * ただ、setValuesには色々制約があるので工夫しています。<br>
 * 
 */

var importCsv = {
    
    /**
     * ファイルから読み込んだバイナリをテキストデータにして返す
     * @param {Object} fileBlob fileBlobオブジェクト
     * @returns {String} stringCode 文字コード sjisなら "Shift_JIS"
     */
    fileToTextData: function(fileBlob, stringCode){
        "use strict";
        var readData = null;
        
        // 文字コード指定読み込みgetDataAsString
        readData = fileBlob.getDataAsString(stringCode);
        if(!readData){
            Logger.log("ERROR:データがありません");
            return null;
        }
        return readData;
    },
    
    /**
     * 対象シートに対象カンマ区切りテキストを１セルずつ書き込む
     * @param {Object} spreadsheetObj 対象になるスプレッドシートオブジェクト
     * @param {Object} sheetObj 対象になるシートオブジェクト
     * @param {String} readData 書き込む文字列
     * @returns {Number} startRow    書き込む最初の行数
     * @returns {Number} startColumn 書き込む最初の列数
     * 
     * setValueで１セルずつ書き込むと超遅いので、setValuesで書き込みます
     */
    writeSheet: function(spreadsheetObj, sheetObj, readData, startRow, startColumn){
        "use strict";
        var csvData = readData.split('\n');         // 行単位分割
        var csvSplit  = csvData[0].split(',');      // 1行目をカンマ単位で分割
        var maxColumn = csvSplit.length;            // 列数
        var maxRow    = csvData.length;             // 行数
        var i, j, iLength, jLength;
        var logStr = "";
        var lineWriteArray = [];
        var srcRowArray = null;
        var range = null;
        var PROGRESS_BODER = 250;                   // 進捗表示用
        
        sheetObj.clear();                   // シートの中身を全クリア
        
        // シートのサイズが小さい場合問題があるので、拡張しておきます
        // setValuesする場合、書き込み先が小さいと落ちる
        utilities.expansionSheetSize(sheetObj, maxColumn + (startColumn-1), maxRow + (startRow-1));
        
        iLength = csvData.length;
        for (i = 0; i < iLength; i++) {
            // 1行文字列を区切り分割して配列化
            srcRowArray = utilities.csvSplit(csvData[i]);
            
            // 区切った配列を列データとして収集
            jLength = srcRowArray.length;
            for (j = 0; j < jLength; j++) {
                lineWriteArray.push(srcRowArray[j]);
            }
            
            // setValuesのセル数限界が251?らしいので、超えない程度で(1行単位で)書き込む
            // [lineWriteArray] → setValuesは二次元配列受付のため
            range = sheetObj.getRange(startRow+i, startColumn, 1, lineWriteArray.length);
            range.setValues( [lineWriteArray] );
            lineWriteArray = [];
            
            // 進捗表示
            if( 0 === (i%PROGRESS_BODER) ){
                logStr = " progress " + i + " / " + iLength;
                Logger.log(logStr);
                spreadsheetObj.toast(logStr);
            }
        }
        return true;
    },

	/**
	 * インポートするファイル選択用ダイアログUI表示
	 * @param {spreadSheetObj}       対象のスプレッドシート
	 * @param {String} titleStr ダイアログのタイトル表示文字列
	 * @param {String} descriptionStr ダイアログの説明文字列
	 * @param {String} butonStr ダイアログのボタン表示文字列
	 *
	 * @example showUiDialog("csvファイルインポート", "インポートするファイルを選んで下さい", "インポート実行");
	 *
	 */ showUiDialog: function (spreadsheetObj, titleStr, descriptionStr, buttonStr) {
        "use strict";
        var app = UiApp.createApplication().setTitle(titleStr);
        var form = app.createFormPanel().setId('frm').setEncoding('multipart/form-data');
        var formContent = app.createVerticalPanel();
        var label1 = app.createLabel("　");
        var label2 = app.createLabel(descriptionStr);
        var file   = app.createFileUpload().setName('csvFileName');
        var button = app.createSubmitButton(buttonStr);
        
        form.add(formContent);
        formContent.add(label1);
        formContent.add(label2);
        formContent.add(file);
        formContent.add(button);
        app.add(form); 
        spreadsheetObj.show(app);
		return true;
	},

	/**
	 * インポート実行
	 * @param {spreadSheetObj}       対象のスプレッドシート
	 * @param {Object} fileBlob      fileBlobオブジェクト
	 * @param {String} stringCode  文字コード sjisなら "Shift_JIS"
	 * @param {String} インポート  e対象になるシート名
	 * @param {Number} startRow    書き込む最初の行数(1 origin)
	 * @param {Number} startColumn 書き込む最初の列数(1 origin)
	 */ run: function (spreadSheetObj, fileBlob, stringCode, sheetNameStr, startRow, startColumn) {"use strict";
        var readTextData      = null;
        if( !fileBlob ){
            var FILE_NAME     = "csv4.txt";
            var forceOnMemory = false;   // シート書き込み強制OFFか？
            readTextData      = this.getTextFile(FILE_NAME);
        }
        else{
        // バイナリ → テキスト変換
          readTextData = this.fileToTextData(fileBlob, stringCode);
          if (!readTextData)
              return false;
        }

		// 対象シートを開く
		var sheetObj = utilities.openSheet(spreadSheetObj, sheetNameStr);
		if (!sheetObj)
			return false;

		// 対象シートに対象カンマ区切りテキストを１セルずつ書き込む
		return this.writeSheet(spreadSheetObj, sheetObj, readTextData, startRow, startColumn);
	}
    
};
