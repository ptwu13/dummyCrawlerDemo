function bulkXml() {
  var sheet = SpreadsheetApp.getActiveSheet();     // 取得現在使用的試算表範圍
  var Num = sheet.getRange("B1").getValue();     // 取得這次爬取的數量，儲存格 B1
  var completedRow = sheet.getRange("H2").getValue();     // 取得已經完成爬取的資料筆數
  var initiateRow = completedRow + 3;     // 將完成爬取的資料筆數+3，作為起始爬取的資料列數
  var totalUrl = sheet.getRange("H1").getValue();     // 取得總資料筆數
  
  for (x=initiateRow;x < Num+initiateRow;x++)  {    // 迴圈執行從起始列開始，直到執行次數 = 本次設定爬取數量
    if (completedRow >= totalUrl) { break; }     // 或是所有的資料都爬完停止執行
    var url = sheet.getRange(x,6).getValue();     // 把爬取的網址，讀進 url 這個變數中
    sheet.getRange(2,9).setValue(url);     // 把 url 變數的值，放到儲存格 I2
    var xpathResult = sheet.getRange("Sheet1!I1:Y1").getValues();     // 把 importxml 的輸出結果，讀到 xpathResult 這個變數中
    sheet.getRange(x,7,1,17).setValues(xpathResult);     //  把 xpathResult 變數的值，放到對應被爬取的 url 儲存格旁邊
    var completedRow = sheet.getRange("H2").getValue();     // 重新取得完成爬取的資料筆數
    SpreadsheetApp.flush();
    }
  
  }


function onOpen(e) {
  SpreadsheetApp.getUi()
      .createMenu('爬取資料')
      .addItem('爬取價格資料', 'bulkXml')
      .addToUi();
}