/**
 * @google spreadsheet用 CSVファイルインポート<br>
 * @author ms32
 * @version 1.0.0
 * 
 * 便利関数集<br>
 * 
 */


var utilities = {
	
	/**
	 * シートサイズの横幅、縦幅を拡張する
	 * 既に指定サイズになっている場合は何も行いません
	 * @param {sheetObj}  対象シート
	 * @param {sizeColumn}     横幅（指定したセル数まで拡張する）
	 * @param {sizeRow}     縦幅（指定したセル数まで拡張する）
	 */ 
	expansionSheetSize: function (sheetObj, sizeColumn, sizeRow) {
        "use strict";
		// 書き込みの高速化のために一括書き込み（setValues）を使用するが、setValuesは
		// シートの行・列数をオーバーして書き込むと落ちるため、予め行と列の足りない分を追加しておく
		var maxColumn = sheetObj.getMaxColumns();
		var maxRow = sheetObj.getMaxRows();
		var addColumnNum = sizeColumn - maxColumn;
		var addRowNum = sizeRow - maxRow;
	
		if (0 < addColumnNum) {
			sheetObj.insertColumns(maxColumn, addColumnNum);
			// 列拡張
		}
		if (0 < addRowNum) {
			sheetObj.insertRows(maxRow, addRowNum);
			// 行拡張
		}
	},
	
	/**
	 * 指定したシート名のシートを返す
	 * 無かったらシートを新規作成して返す
	 * @param {spreadSheetObj} 対象のスプレッドシート
	 * @param {string} 対象のシート名
	 */ 
	openSheet: function (spreadSheetObj, sheetNameStr) {
        "use strict";
		var sheetObj = spreadSheetObj.getSheetByName(sheetNameStr);
	
		if (null === sheetObj) {
			Logger.log("対象シートがシートが見つかりません [" + sheetNameStr + "]");
	
			// 指定したシート名のシートが見つからなかった場合は、新規作成する
			var temp = spreadSheetObj.getSheetByName("fl_template");
			sheetObj = spreadSheetObj.insertSheet(sheetNameStr, 1, {
				template : temp
			});
	
			if (null === sheetObj)
				Logger.log("エラー:新規にシート追加できませんでした [" + sheetNameStr + "]");
			else
				Logger.log("新規にシートを追加しました [" + sheetNameStr + "]");
		}
		return sheetObj;
	},

    /**
    * すべての文字列 s1をs2に置き換える
    * ( String.replaceでは最初の1文字しか置き換えられない)
    * @param {String}  ベース文字列
    * @param {src}     置き換え元文字
    * @param {dest}    置き換え先文字
    * @return {XXX}
    */
    replaceAll: function(expression, src, dest){
        "use strict";    /* 変数宣言強制 */
        return expression.split(src).join(dest);
    },


    /**
     * テキストデータ１行をカンマ単位に分離する
     * ダブルクォーテーション間のカンマなどの例外にも対応する
     * @param {textLine}  
     * @return {カンマ単位に分離した文字列配列}
     */
    csvSplit: function(lineStr) {
        "use strict";    
        var result   = null;
        var tmpArray = null;
        var i        = 0;
        var backup   = lineStr;
        
        tmpArray = lineStr.split('"');
        if( 1 < tmpArray.length ){
            lineStr = "";
            tmpArray[1] = this.replaceAll(tmpArray[1], ",", " ");
    
            for(i = 0; i < tmpArray.length; i++) lineStr += tmpArray[i];
        }
        result = lineStr.split(',');
        return result;    
    },
    
    /**
     * googleドライブ上にあるテキストファイルを指定して読み込む
     * @param {String} readFileName 読み込むテキストファイル名  
     * @return {contents} 読み込み後のテキストデータ
     */
    getTextFile: function (readFileName) {
        "use strict";
        var files      = null;
        var fileBlob   = null;
        var contents   = null;
        var i          = 0;
        var fileName   = "";
        
        files = DocsList.getFiles();
        for(i = 0; i < files.length; i++){
            fileName = files[i].getName();
            if(readFileName === fileName){
                fileBlob = files[i].getBlob();
                contents = fileBlob.getDataAsString("Shift_JIS");
                return contents;
            }
        }
        return null;
    }


    
};
