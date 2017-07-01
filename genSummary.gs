// header enum for representing row index
var rearrangedheadEnum = { DATE        : 0, // TBD
                           DATE_COL    : "A",
                           AMOUNT      : 1,  // TBD
                           AMOUNT_COL  : "B",
                           ACTION      : 2, // TBD
                           ACTION_COL  : "C", // TBD
                           PRICE       : 3, // TBD
                           QUANTITY    : 4, // TBD	
                           QUANTITY_COL: "E", // TBD
                           SYMBOL	   : 5,
                           SYMBOL_COL  : "F",
                           COMMISSION  : 6 // TBD
                          };

var rearrangedheadContent = 
                  [  "DATE",
                     "AMOUNT",
                     "ACTION",	
                     "PRICE",
                     "QUANTITY",
                     "SYMBOL",
                     "COMMISSION"
                  ];
 
var summaryheadEnum = { SYMBOL_COL        : "A", // TBD
                        POSITIONs_COL     : "B",  // TBD
                        LASTPRICE_COL     : "D",
                        preMarketValue_COL : "D",
                        MarketValue_COL   : "E"
                          };

var summaryheadContent = ["SYMBOL", "POSITIONs", "AvgCOST", "LastPrice", "MarketValue"]

function onOpen(e) {
  var menu = SpreadsheetApp.getUi().createAddonMenu(); // Or DocumentApp or FormApp.
  menu.addItem('Gernerate summary', 'genSummary');
  menu.addToUi();
}

function genSummary() {
  var newData = [];
  var datasheetname = "rearranged_data";
  var summarysheetname = "US_investment_summary";
  
  // parse and extract raw data from TD history sheet
  newData = _parseTD(newData, "TD_history");
  newData = _parseSCHWAB(newData, "SCHWAB_history");
  newData = _parseSubBrokerage(newData, "subbrokerage_history");
  
  // update the extracted data to sheet
  _updateDatasheet(newData, datasheetname, summarysheetname);
  
  
  // extract unique stock symbol
  symbolTable = _extractSymbols(newData);
  //
  _updateSummarysheet(symbolTable, newData.length, summarysheetname, datasheetname);

}

function _parseTD(newData, sheetName) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (sheet != null){
     var rawData = sheet.getDataRange().getValues();
     // header enum for representing row index
     var headEnum = { DATE        : 0,
                      DESCRIPTION : 2,
                      QUANTITY    : 3,	
                      SYMBOL	  : 4,
                      PRICE       : 5,
                      COMMISSION  : 6,
                      AMOUNT      : 7
                    };
    
     for (i in rawData) {
       var row = rawData[i];
       var action = row[headEnum.DESCRIPTION].toString().split(" ");
       var actionDescription = " ";
       // BUY
       if (action[0] == "Bought") {
         actionDescription = "BUY";
       }
       // SELL
       else if (action[0] == "Sold") {
         actionDescription = "SELL";
       }
       // DIVIDEND
       else if (action[0] == "ORDINARY" && action[1] == "DIVIDEND") {
         actionDescription = "DIVIDEND";
       }
       else if (action[0] == "W-8" && action[1] == "WITHHOLDING") {
         actionDescription = "WITHHOLDING";
       }
       else {
         continue;
       }
                
       // insert new data row to newData
       newData.push([new Date(row[headEnum.DATE]),
                     row[headEnum.AMOUNT],
                     actionDescription,
                     row[headEnum.PRICE],
                     row[headEnum.QUANTITY],
                     row[headEnum.SYMBOL],
                     row[headEnum.COMMISSION]
                     ]);
     }

     return newData; 
  }
}

function _parseSCHWAB(newData, sheetName) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (sheet != null){
     var rawData = sheet.getDataRange().getValues();
     // header enum for representing row index
     var headEnum = { DATE        : 0,
                      ACTION      : 1,
                      SYMBOL      : 2,	
                      DESCRIPTION : 3,
                      QUANTITY    : 4,
                      PRICE       : 5,
                      COMMISSION  : 6,
                      AMOUNT      : 7
                    };
    
     for (i in rawData) {
       var row = rawData[i];
       var action = row[headEnum.ACTION].toString().split(" ");
       var actionDescription = " ";
       // BUY
       if (action[0] == "Buy" || (action[0] == "Reinvest" && action[1] == "Shares")) {
         actionDescription = "BUY";
       }
       // SELL
       else if (action[0] == "Sell") {
         actionDescription = "SELL";
       }
       // DIVIDEND
       else if (action[0] == "Reinvest" && action[1] == "Dividend") {
         actionDescription = "DIVIDEND";
       }
       else if (action[0] == "NRA" && action[1] == "Tax") {
         actionDescription = "WITHHOLDING";
       }
       else {
         continue;
       }
                
       // insert new data row to newData
       newData.push([new Date(row[headEnum.DATE]),
                     row[headEnum.AMOUNT],
                     actionDescription,
                     row[headEnum.PRICE],
                     row[headEnum.QUANTITY],
                     row[headEnum.SYMBOL],
                     row[headEnum.COMMISSION]
                     ]);
     }

     return newData; 
  }
}

function _parseSubBrokerage(newData, sheetName) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (sheet != null){
     var rawData = sheet.getDataRange().getValues();
     // header enum for representing row index
     var headEnum = { DATE        : 0,
                      DESCRIPTION : 2,
                      QUANTITY    : 3,	
                      SYMBOL	  : 4,
                      PRICE       : 5,
                      COMMISSION  : 6,
                      AMOUNT      : 7
                    };
    
     for (i in rawData) {
       var row = rawData[i];
       var action = row[headEnum.DESCRIPTION].toString().split(" ");
       var actionDescription = " ";
       // BUY
       if (action[0] == "Bought") {
         actionDescription = "BUY";
       }
       // SELL
       else if (action[0] == "Sold") {
         actionDescription = "SELL";
       }
       // DIVIDEND
       else if (action[0] == "ORDINARY" && action[1] == "DIVIDEND") {
         actionDescription = "DIVIDEND";
       }
       else if (action[0] == "W-8" && action[1] == "WITHHOLDING") {
         actionDescription = "WITHHOLDING";
       }
       else {
         continue;
       }
                
       // insert new data row to newData
       newData.push([new Date(row[headEnum.DATE]),
                     row[headEnum.AMOUNT],
                     actionDescription,
                     row[headEnum.PRICE],
                     row[headEnum.QUANTITY],
                     row[headEnum.SYMBOL],
                     row[headEnum.COMMISSION]
                     ]);
     }

     return newData; 
  }
}

function _updateDatasheet(newData, datasheetname, summarysheetname)
{
     //Logger.log(testDate);
     var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(datasheetname);
     sheet = sheet.clear();
     // update header
     sheet.getRange(1, 1, 1, rearrangedheadContent.length).setValues([rearrangedheadContent]);
     // update contents
     sheet.getRange(2, 1, newData.length, newData[0].length).setValues(newData);
     // update last row with total market value for xirr
     sheet.getRange(newData.length+2, 1, 1, 2).setFormulas([["=Today()", _lookupTotalmarketvalueFormula(summarysheetname)]]);
}

function _lookupTotalmarketvalueFormula(sheetname) {
  var formula = "=VLOOKUP(\"Total market value\","+ sheetname + "!" + summaryheadEnum.preMarketValue_COL + ":" + summaryheadEnum.MarketValue_COL + ",2,false)";
  return formula;
}
                                                          
function _extractSymbols(dataArr) {
  var symbolTable = {}; 
  for (i in dataArr){
    var row = dataArr[i];
    var sym = row[rearrangedheadEnum.SYMBOL];
    symbolTable[sym] = sym; 
  }
  return symbolTable;
}

function _updateSummarysheet(symbolMap, datalength, summarySheetname, dataSheetname) {
  var sheet2 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(summarySheetname);
  sheet2 = sheet2.clear();
  
  // update header
  sheet2.getRange(1, 1, 1, 5).setValues([summaryheadContent]);
  // update summary table
  var formulaArr =  [];
  var datalastrowIdx = datalength + 1;
  var idx = 1;
  for (s in symbolMap) {
     // processing formulas
     var currowIdx = idx + 1;
     var symbolPos = 'A' + currowIdx.toString();
     var positionPos = 'B' + currowIdx.toString();
     var lastpricePos = 'D' + currowIdx.toString(); 
     formulaArr.push(  [ _getSymbolFormula(s),
                         _getPostionFormula(2, datalastrowIdx, currowIdx, "rearranged_data"), // positions
                         _getAvgCostFormula(2, datalastrowIdx, currowIdx, "rearranged_data"), // avgcost
                         _getlastValueFormula(currowIdx), // LastPrice
                         _getmarketValueFormula(currowIdx) // marketValue
                       ]);                   
     idx += 1;
  }
  
  // add total market value
  formulaArr.push([ "=\"\"", "=\"\"", "=\"\"", "=\"Total market value\"", _getTotalmarketValueFormula(2, formulaArr.length+1, summaryheadEnum.MarketValue_COL) ]);
  
  // add total divident w/ withholding
  formulaArr.push([ "=\"\"", "=\"\"", "=\"\"", "=\"Total dividend with withholding\"", _getTotaldividendFormula(2, datalength+1, dataSheetname) ]);
  
  // add xirr calculation
  formulaArr.push([ "=\"\"", "=\"\"", "=\"\"", "=\"XIRR\"", _getXIRRFormula(2, datalength+2, rearrangedheadEnum.AMOUNT_COL, rearrangedheadEnum.DATE_COL, dataSheetname) ]);
  
  // update the rest table with formulas
  sheet2.getRange(2, 1, formulaArr.length, formulaArr[0].length).setFormulas(formulaArr);
  
  // update xirr format to percentage
  var xirrCellPos = summaryheadEnum.MarketValue_COL + String(formulaArr.length+1);
  sheet2.getRange(xirrCellPos).setNumberFormat("###.00%");
  
} 

function _getPostionFormula(datafirstrow, datalastrow, currow, datasheetname) {
  // SUMIFS(sum_range, criteria_range1, criteria1, [criteria_range2, criteria2], ...)
  var summarySymbolPos = summaryheadEnum.SYMBOL_COL + currow.toString();
  var dataQuantityRange = rearrangedheadEnum.QUANTITY_COL + datafirstrow.toString() + ":" + rearrangedheadEnum.QUANTITY_COL + datalastrow.toString();
  var dataActionRange = rearrangedheadEnum.ACTION_COL   + datafirstrow.toString() + ":" + rearrangedheadEnum.ACTION_COL + datalastrow.toString();
  var dataSymbolRange = rearrangedheadEnum.SYMBOL_COL   + datafirstrow.toString() + ":" + rearrangedheadEnum.SYMBOL_COL + datalastrow.toString();
  var formula = "SUMIFS(" + 
                 datasheetname + "!" + dataQuantityRange + "," +
                 datasheetname + "!" + dataActionRange + "," +
                 '\"BUY\",' +
                 datasheetname + "!" + dataSymbolRange + "," +
                 summarySymbolPos + ')' +
                 "-" +
                 "SUMIFS(" + 
                 datasheetname + "!" + dataQuantityRange + "," +
                 datasheetname + "!" + dataActionRange + "," +
                 '\"SELL\",' +
                 datasheetname + "!" + dataSymbolRange + "," +
                 summarySymbolPos + ')';
  return  formula;
}

function _getSymbolFormula(symbol) {
  var formula =  "=\"" + symbol + "\"";
  return formula;
}
function _getAvgCostFormula(datafirstrow, datalastrow, currow, datasheetname) {
   // SUMIFS(sum_range, criteria_range1, criteria1, [criteria_range2, criteria2], ...)
  var summarySymbolPos = summaryheadEnum.SYMBOL_COL + currow.toString();
  var summaryPositionPos = summaryheadEnum.POSITIONs_COL + currow.toString();
  var dataAmountRange = rearrangedheadEnum.AMOUNT_COL + datafirstrow.toString() + ":" + rearrangedheadEnum.AMOUNT_COL + datalastrow.toString();
  var dataSymbolRange = rearrangedheadEnum.SYMBOL_COL   + datafirstrow.toString() + ":" + rearrangedheadEnum.SYMBOL_COL + datalastrow.toString();
  var formula = "ABS(SUMIFS(" + 
                 datasheetname + "!" + dataAmountRange + "," +
                 datasheetname + "!" + dataSymbolRange + "," +
                 summarySymbolPos + '))' +
                 "/" +
                 summaryPositionPos;
  return  formula;
}

function _getlastValueFormula(currow) {
  var summarySymbolPos = summaryheadEnum.SYMBOL_COL + currow.toString();
  var formula = "GoogleFinance(" + summarySymbolPos + ",\"price\")";
  return formula;
}

function _getmarketValueFormula(currow) {
  var summaryPositionPos = summaryheadEnum.POSITIONs_COL + currow.toString();
  var summaryLastpricePos = summaryheadEnum.LASTPRICE_COL + currow.toString();
  var formula = summaryPositionPos + "*" + summaryLastpricePos;
  return formula;
}

function _getTotalmarketValueFormula(firstrow, lastrow, marketValCol) {
  var marketValueRange = marketValCol + firstrow.toString() + ':' + marketValCol + lastrow.toString();
  var formula =  "SUM(" + marketValueRange + ")";
  return formula;
}
  
function _getXIRRFormula(firstrow, lastrow, amountCOL, dateCOL, sheetname) {
  var amountRange = sheetname + "!" + amountCOL + firstrow.toString() + ":" + amountCOL + lastrow.toString(); 
  var dateRange   = sheetname + "!" + dateCOL + firstrow.toString() + ":" + dateCOL + lastrow.toString();
  var formula = "XIRR(" + amountRange + "," + dateRange + ",false)";
  return formula;
}

function _getTotaldividendFormula(datafirstrow, datalastrow, datasheetname) {
  // SUMIFS(sum_range, criteria_range1, criteria1, [criteria_range2, criteria2], ...)
  var dataAmountRange = rearrangedheadEnum.AMOUNT_COL + datafirstrow.toString() + ":" + rearrangedheadEnum.AMOUNT_COL + datalastrow.toString();
  var dataActionRange = rearrangedheadEnum.ACTION_COL   + datafirstrow.toString() + ":" + rearrangedheadEnum.ACTION_COL + datalastrow.toString();
  var formula = "SUMIFS(" + 
                 datasheetname + "!" + dataAmountRange + "," +
                 datasheetname + "!" + dataActionRange + "," +
                 '\"DIVIDEND\")' +
                 "-" +
                 _getTotalwithholdingFormula(datafirstrow, datalastrow, datasheetname);
  return formula;
}

function _getTotalwithholdingFormula(datafirstrow, datalastrow, datasheetname) {
  // SUMIFS(sum_range, criteria_range1, criteria1, [criteria_range2, criteria2], ...)
  var dataAmountRange = rearrangedheadEnum.AMOUNT_COL + datafirstrow.toString() + ":" + rearrangedheadEnum.AMOUNT_COL + datalastrow.toString();
  var dataActionRange = rearrangedheadEnum.ACTION_COL   + datafirstrow.toString() + ":" + rearrangedheadEnum.ACTION_COL + datalastrow.toString();
  var formula = "ABS(SUMIFS(" + 
                 datasheetname + "!" + dataAmountRange + "," +
                 datasheetname + "!" + dataActionRange + "," +
                 '\"WITHHOLDING\"))';
  return formula;
}
