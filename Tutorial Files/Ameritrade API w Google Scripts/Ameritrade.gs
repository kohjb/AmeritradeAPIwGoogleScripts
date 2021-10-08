function amtd_ShowPane() {
//Open a SidePane asynchronously. The html will return by calling the function amtd_backfromPane

  linkURL = "https://auth.tdameritrade.com/auth?response_type=code&redirect_uri=https%3A%2F%2F127.0.0.1&client_id="+ amtd.apikey +"%40AMER.OAUTHAP";
  var html = HtmlService.createTemplateFromFile('amtd_SidePane')
    .evaluate();
  SpreadsheetApp.getUi().showSidebar(html);
}

function amtd_backfromPane(d) {
//Called after user clicks Step 2 button on SidePane, return here with dictionary d

  //amtd_GetTokens(d.returnURI);
  
}

function amtd_GetQuote(stockSymbol) {
  return amtd.amtd_GetQuote(stockSymbol);
}

function amtd_GetOptionPrice(stockSymbol, contractType, strike, expDate, priceType) {
  return amtd.amtd_GetOptionPrice(stockSymbol, contractType, strike, expDate, priceType);
}

function amtd_Test() {
  return amtd.amtd_Test();
}