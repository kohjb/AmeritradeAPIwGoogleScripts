var userProperties = PropertiesService.getUserProperties();

//******************************MAIN FUNCTIONS*****************************************************************************************
function test() {

//  slog(amtd_GetQuote("INTC"));
//  slog(amtd_GetOptionChain("INTC", "Call")

  var s = "{name=22,1450=[{totalVolume=67, symbol=GOOG_022120C1450, openInterest=1624, optionDeliverablesList=null, delta=0.579, description=GOOG Feb 21 2020 1450 Call, openPrice=0, volatility=24.581, timeValue=36.76, theta=-0.838, lastTradingDay=1.5822612E12, lowPrice=34.8, highPrice=42.6, askSize=1, theoreticalVolatility=29, expirationDate=1.5823368E12, markPercentChange=-3.24, netChange=0, settlementType= , isIndexOption=null, percentChange=0, expirationType=R, last=42.6, mini=false, bidSize=1, multiplier=100, daysToExpiration=21, inTheMoney=true, tradeTimeInLong=1.580417392981E12, tradeDate=null, putCall=CALL, quoteTimeInLong=1.580417999968E12, markChange=-1.44, lastSize=0, nonStandard=false, ask=43.4, rho=0.484, exchangeName=OPR, deliverableNote=, closePrice=44.39, bid=42.5, bidAskSize=1X1, theoreticalOptionValue=42.95, mark=42.95, gamma=0.004, vega=1.404, strikePrice=1450}]}";
  var json = s;
  slog(json["name"]);

}

/**
 * Call the Ameritrade API to get the quote.
 * @param {string} stockSymbol The symbol of the stock to look up
 * @return {number} The current price of the stock
 * @customfunction
  */
function amtd_GetQuote(stockSymbol) {
  // Call the Ameritrade API to get the quote. 
  return amtd.amtd_GetQuote(stockSymbol);
}

/**
 * Call Ameritrade API to get the closing prices of one or more stockSymbols.
 *
 * @param {string} the stock's symbol - comma seperated list of symbols, e.g. "AAPL,/EQ"
 * @return {string} the closing price - comma seperated list of prices, e.g. "167.30,4369.50"
 * @customfunction
  */
function amtd_GetQuotes(stockSymbols) {
  // Call the Ameritrade API to get the quote. 
  return amtd.amtd_GetQuotes(stockSymbols);
}

/**
 * Call the Ameritrade API to get the Option's Price.
 * @param {string} stockSymbol The symbol of the stock to look up
 * @param {string} contractType "PUT" or "CALL" option
 * @param {number} strike The Strike Price of the contract
 * @param {Date} expDate The Expiry Date of the contract
 * @param {string} priceType "B" or "W" representing Buy or Write/Sell Price of the contract
 * @return {number} The current price of the stock 
 * @customfunction
 */
function amtd_GetOptionPrice(stockSymbol, contractType, strike, expDate, priceType) {
  //Call Ameritrade API and get the Option's price
  return amtd.amtd_GetOptionPrice(stockSymbol, contractType, strike, expDate, priceType);
}

/**
 * Call the Ameritrade API to get the price history of a stock.
 *
 * @param {string} stockSymbol the stock's symbol
 * @param {date} stockDate the expiry date of the contract
 * @return {number} the option's price
 * @customfunction
 */
function amtd_GetPriceHistory(stockSymbol, stockDate) {
  //Call the Ameritrade API to get the price history
  return amtd.amtd_GetPriceHistory(stockSymbol, stockDate);
}

//*****************************AUTHENTICATION FUNCTIONS****************************************************************
function amtd_showPane(){
//Open a SidePane asynchronously. The html will return by calling the function amtd_backfromPane
//

  linkURL = "https://auth.tdameritrade.com/auth?response_type=code&redirect_uri=https%3A%2F%2F127.0.0.1&client_id="
    + amtd.apikey + "%40AMER.OAUTHAP";
  var html = HtmlService.createTemplateFromFile('amtd_SidePane')  //.createHtmlOutputFromFile('SidePane')
      .evaluate();
    
  SpreadsheetApp.getUi() 
      .showSidebar(html);
}

function amtd_backfromPane(d) {  
//After user clicks Enter button on SidePane, return here with dictionary d. 

//  slog("Returned URI: " + d.returnURI);
  amtd.amtd_GetTokens(d.returnURI);
}

function amtd_getAccessToken() {
  return amtd.amtd_getAccessToken()
}

function amtd_getAccessTokenTime() {
  var s = amtd.amtd_getAccessTokenTime()
  var d = new Date(s)
  return d
}

function amtd_getRefreshToken() {
  return amtd.amtd_getRefreshToken()
}

function amtd_getRefreshTokenTime() {
  var s = amtd.amtd_getRefreshTokenTime()
  var d = new Date(s)
  return d
}


function showSignonLink(linkURL) {
  var html = '<html><body>Log in to Ameritrade to authenticate your account: Click <a href="'+linkURL+'" target="blank" onclick="google.script.host.close()">'+"here"+'</a></body></html>';
  var ui = HtmlService.createHtmlOutput(html)
  SpreadsheetApp.getUi().showModalDialog(ui,"Log in to Ameritrade");
}



























