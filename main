var ss = SpreadsheetApp.getActiveSpreadsheet();
var purchaseTab = ss.getSheetByName("myPurchases");
var calculationTab = ss.getSheetByName("Calculation");


function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Portfolio Analysis')
      .addItem('Run Analysis', 'PortfolioAnalysis')
      .addToUi();
}

function PortfolioAnalysis(){
  var userInput = SelectData(5, 5, purchaseTab);
  var allData = SelectData(5, purchaseTab.getLastColumn(), purchaseTab);
  var i = 5; //First row with data
  userInput.forEach(
    function(row){
      var purchaseDate = ConvertToEPOCH(row[0]); // Need to make this not EPOCH and then convert it where it needs it.
      var userTicker = row[2];
      var shares = row[3];
      var costBasis = row[4];
      SetStockName(i, userTicker, purchaseTab);
      var totalCost = CalculateTotalCost(i, shares, costBasis, purchaseTab);
      var capitalGains = parseFloat(CalculateCapitalGains(i, userTicker, shares, totalCost, purchaseTab));
      var dividendGains = parseFloat(CalculateDividend(i, userTicker, purchaseDate, shares));
      var lastDividendPayment_Date = 0;
      var totalGains = capitalGains + dividendGains;
      SetTotalGain(i, totalGains, totalCost, purchaseTab); 
      //Logger.log(capitalGains, dividendGains, userTicker);
      i = i+1;
    }
  ) //End of the ForEach Loop
}

function SetTotalGain(row, gain, cost, tab){
  var totalGainCell_dollar = tab.getRange(row,8);
  var totalGainCell_percent = tab.getRange(row, 9);
  totalGainCell_dollar.setValue(gain);
  totalGainCell_percent.setValue(gain/cost);
  
}

function CalculateDividend(row, ticker, purchaseDate, shares){
  var todayDate = GetTodayDate();
  var dividendURL = MakeDividendURL(ticker, purchaseDate, todayDate);
  var cellFunction = '=INDEX(IMPORTHTML("'+dividendURL+'","table",1),0,2)';
  calculationTab.getRange("A1").setValue(cellFunction);
  var dividendRows = calculationTab.getDataRange();
  var firstRow = 2;
  var lastRow = dividendRows.getLastRow() - firstRow +1;
  var dividendPayments = SUM(CleanDividendData(calculationTab.getRange(firstRow, 1, lastRow, 1).getValues()));
  //Logger.log(dividendPayments);
  return (shares*dividendPayments).toFixed(2);
}

function SUM(data){
  var i;
  var sum = 0;
  for (i = 0; i < data.length; i++){
    sum = sum + Number(data[i]);
  }
  return sum.toFixed(2);
}

function CleanDividendData(data){
  var paymentsOnly = []
  data.forEach(
    function(row){
      paymentsOnly.push(row[0].split("*")[1]);
    }
  )
  return paymentsOnly
}    

function MakeDividendURL(ticker, purchaseDate, todayDate){
    var url = "https://finance.yahoo.com/quote/"+ticker+"/history?period1="+purchaseDate+"&period2="+todayDate+"&interval=div%7Csplit&filter=div&frequency=1d&includeAdjustedClose=true";
    return url;
}

function ConvertToEPOCH(timeData){
  var epochTime = timeData.getTime()/1000;
  return epochTime;
}

function GetTodayDate(){
  var dateObject = new Date();
  var today = dateObject.getTime()/1000;
  return Math.trunc(today);
}

function CalculateCapitalGains(row, ticker, shares, totalCost, purchaseTab){
  var currentValue = GetCurrentValue(row, ticker, purchaseTab);
  return (currentValue*shares)-totalCost;
}

function GetCurrentValue(row, ticker, tab){
  var currentValueCell = tab.getRange(row, 7);
  if (currentValueCell.getValue() == ''){
    var currentPrice = '=GOOGLEFINANCE("'+ticker+'","price")';
    currentValueCell.setValue(currentPrice);
  }
  return currentValueCell.getValue();
}


function CalculateTotalCost(row, shares, costBasis, tab){
  var totalCostCell = tab.getRange(row,6);
  if (totalCostCell.getValue() == ''){
    var totalCost = shares * costBasis;
    totalCostCell.setValue(totalCost);
  }
  return totalCostCell.getValue();
}

function SetStockName(row, ticker, tab){
  var companyNameCell = tab.getRange(row,2);
  if (companyNameCell.getValue() == ''){
    var companyName = '=GOOGLEFINANCE("'+ticker+'","name")';
    companyNameCell.setValue(companyName);
  }
}


function SelectData(firstRowWithData, lastColumnWithData, tab){
  var rows = tab.getDataRange();
  var lastRow = rows.getLastRow() - firstRowWithData + 1;
  var data = tab.getRange(firstRowWithData, 1, lastRow, lastColumnWithData).getValues();
  return data;
}
