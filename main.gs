/**
 * @google spreadsheet用 CSVファイルインポート<br>
 * @author ms32
 * @version 1.0.0
 * 
 * google apps script スプレッドシートから直接呼ばれるコールバック関連
 * 
 * 
 */


function onOpen() {
	"use strict";
    var spreadsheetObj = SpreadsheetApp.getActiveSpreadsheet();
    var menuList       = [];
    menuList.push({
        name : "CSVファイルインポート",
        functionName : "importCSV"
    });
    spreadsheetObj.addMenu("csvインポート", menuList);
}

function importCSV() {
	"use strict";
    var spreadsheetObj = SpreadsheetApp.getActiveSpreadsheet();
	var titleStr		=	"csvファイルインポート";
	var descriptionStr	=	"インポートするファイルを選んで下さい";
	var buttonStr		=	"インポート実行";
	importCsv.showUiDialog(spreadsheetObj, titleStr, descriptionStr, buttonStr);
}

function doPost(e) {
	"use strict";
    var spreadsheetObj		= SpreadsheetApp.getActiveSpreadsheet();
	var startRow			= 3;			// 書き込み開始行数(1 origin)
	var startColumn			= 1;			// 書き込み開始列数(1 origin)
	var STRING_CODE			= "Shift_JIS";	// CSV文字コード指定
	var EXPORT_SHEET_NAME	= "export";		// 出力シート名
	var logStr              = "";
	var fileBlob			= e.parameter.csvFileName;
    var app                 = UiApp.getActiveApplication();
    
	importCsv.run(spreadsheetObj, fileBlob, STRING_CODE, EXPORT_SHEET_NAME, startRow, startColumn);
	
    logStr = "import finish";
    Logger.log(logStr);
    spreadsheetObj.toast(logStr);
    return app.close();
}


function test() {
    "use strict";
    var spreadsheetObj = SpreadsheetApp.getActiveSpreadsheet();
    var startRow            = 3;            // 書き込み開始行数
    var startColumn         = 1;            // 書き込み開始列数
    var STRING_CODE         = "Shift_JIS";  // CSV文字コード指定
    var EXPORT_SHEET_NAME   = "export";     // 出力シート名
    
    
    importCsv.run(spreadsheetObj, null, STRING_CODE, EXPORT_SHEET_NAME, startRow, startColumn);
}

