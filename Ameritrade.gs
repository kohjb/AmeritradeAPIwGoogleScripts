
var apikey = "YourAPIKeyHere....";
var userProperties = PropertiesService.getUserProperties();

function amtd_ShowPane() {
//Open a SidePane asynchronously. The html will return by calling the function amtd_backfromPane

  linkURL = "https://auth.tdameritrade.com/auth?response_type=code&redirect_uri=https%3A%2F%2F127.0.0.1&client_id="+ amtd.apikey +"%40AMER.OAUTHAP";
  var html = HtmlService.createTemplateFromFile('amtd_SidePane')
    .evaluate();
  SpreadsheetApp.getUi().showSidebar(html);
}

function amtd_backfromPane(d) {
//Called after user clicks Step 2 button on SidePane, return here with dictionary d

  amtd_GetTokens(d.returnURI);
  
}

//******************************MAIN FUNCTIONS*****************************************************************************************

/**
 * Call Ameritrade API to get the closing price of stockSymbol.
 *
 * @param {string} the stock's symbol
 * @return {number} the closing price
 */
 function amtd_GetQuote(stockSymbol) {

  var authorization = amtd_GetBearerString_();
  var options = {
    "method" : "GET",
    "headers" :  {"Authorization" : authorization},
    "apikey" : apikey
  }
  var myurl="https://api.tdameritrade.com/v1/marketdata/"+ stockSymbol +"/quotes";
  var result=UrlFetchApp.fetch(myurl, options);
  
  //Parse JSON
  var contents = result.getContentText();
  var json = JSON.parse(contents);
  var stock = json[stockSymbol];
  
  closePrice = stock["closePrice"];
  
  return closePrice;
}

/**
 * Call Ameritrade API and get the Option's price.
 *
 * @param {string} the stock's symbol
 * @param {string} CALL or PUT contract
 * @param {number} the strike price of the contract
 * @param {date} the expiry date of the contract
 * @param {string} bid or ask price
 * @return {number} the option's price
 */
function amtd_GetOptionPrice(stockSymbol, contractType, strike, expDate, priceType) {
//Call Ameritrade API and get the Option's price

  //For debugging
//  var stockSymbol = "GOOG";
//  var contractType = "PUT";
//  var strike = 1450;
//  var expDate = "2020-02-21";
//  var priceType = "bid";

  expDate = Utilities.formatDate(expDate, "GMT+8", "yyyy-MM-dd").toString();
  var authorization = amtd_GetBearerString_();
  var options = {
    "method" : "GET",
    "headers" :  {"Authorization" : authorization},
    "apikey" : apikey
  }
  var foptions = "?symbol="+stockSymbol+"&contractType="+contractType+"&strike="+strike+"&fromDate="+expDate+"&toDate="+expDate;
  
  var myurl="https://api.tdameritrade.com/v1/marketdata/chains"+foptions;

  var result=UrlFetchApp.fetch(myurl, options);
  
  //Parse JSON
  //Sample JSON: 
    //{"symbol":"GOOG","status":"SUCCESS","underlying":null,"strategy":"SINGLE","interval":0,"isDelayed":true,"isIndex":false,"interestRate":2.42788,"underlyingPrice":1436.045,
    //"volatility":29,"daysToExpiration":0,"numberOfContracts":1,"callExpDateMap":{},"putExpDateMap":{"2020-03-20:49":{"1450":[{"putCall":"PUT","symbol":"GOOG_032020P1450",
    //"description":"GOOG Mar 20 2020 1450 Put","exchangeName":"OPR","bid":62,"ask":63.9,"last":46.6,"mark":62.95,"bidSize":3,"askSize":31,"bidAskSize":"3X31","lastSize":0,
    //"highPrice":46.6,"lowPrice":46.6,"openPrice":0,"closePrice":45.42,"totalVolume":1,"tradeDate":null,"tradeTimeInLong":1580481504947,"quoteTimeInLong":1580488460635,
    //"netChange":1.18,"volatility":27.338,"delta":-0.508,"gamma":0.003,"theta":-0.531,"vega":2.118,"rho":-0.994,"openInterest":193,"timeValue":32.41,"theoreticalOptionValue":62.95,
    //"theoreticalVolatility":29,"optionDeliverablesList":null,"strikePrice":1450,"expirationDate":1584752400000,"daysToExpiration":49,"expirationType":"R","lastTradingDay":1584676800000,
    //"multiplier":100,"settlementType":" ","deliverableNote":"","isIndexOption":null,"percentChange":2.59,"markChange":17.53,"markPercentChange":38.58,"inTheMoney":true,"mini":false,"nonStandard":false}]}}}
  var contents = result.getContentText();
  var json = JSON.parse(contents);
  
  var status = json["status"];
  if (status == "SUCCESS") {
    mapname = contractType.toLowerCase() + "ExpDateMap";
    
    putcallmap = json[mapname];
    
    for(var k1 in putcallmap) if (k1.indexOf(expDate)>-1) {
      datemap = putcallmap[k1];
      for(var k2 in datemap) if (k2.indexOf(strike.toString())>-1) {
        pricemap = datemap[k2];
        price = pricemap[0];
        optprice = price[priceType];
        return  optprice;
      }
    } else {
      return 0;
    }
  }
}

function amtd_GetPriceHistory(stockSymbol, stockDate) {  
  // Call the Ameritrade API to get the price history. 
  // stockDate  is at 6pm
  // Returns a JSON object. E.g.
    //{"candles":[{"open":320.93,"high":322.68,"low":308.29,"close":309.51,"volume":49897096,"datetime":1580450400000}],"symbol":"AAPL","empty":false}

  //For debugging
//  var stockSymbol = "AAPL"
//  var stockDate = "1580364000000"; //epochdate("1/29/2020"), 
  
  //Setup key variables
  stockDate = epochdate_(stockDate);    //Change Date object to epochdate. e.g. 30/1/20 = 1580364000000
  var endDate = (Number(stockDate)+57700000).toString();  //Set closing time to a bit past 4pm
  var startDate = (Number(stockDate)+57600000).toString();    //Set starting time to 4pm
  
  var authorization = amtd_GetBearerString_();
  
  var options = {
    "method" : "GET",
    "headers": {"Authorization": authorization},
    "apikey" : apikey
  }

  var periodType = "month";    //day, month, year, or ytd (year to date). Default is day.
  var period = 1;             //depends on periodType. If month, 1,2,3,6. 
  var frequencyType = "daily";    //depends on periodType. If month, daily, weekly
  var frequency = 1;    //depends on frequencyType. If daily, 1.
  var dateinfoType  = "close";    //open, high, low, close, volume, datetime
  var needExtendedHoursData = "false";    //true to get ext hours.
    
  var foptions = "?periodType="+periodType+"&period="+period+"&frequencyType="+frequencyType+"&frequency="+frequency+"&endDate="+endDate+"&startDate="+startDate+"&needExtendedHoursData="+needExtendedHoursData;
  
  //Create url and fetch it
  var myurl="https://api.tdameritrade.com/v1/marketdata/"+stockSymbol+"/pricehistory"+foptions;
  var result=UrlFetchApp.fetch(myurl, options);  

  //Parse JSON
  var contents = result.getContentText();
  var json = JSON.parse(contents);
  var candles  = json["candles"];
  var dateinfo = candles[0];     
  var price = dateinfo[dateinfoType];
  return(price);
}

//*****************************AUTHENTICATION FUNCTIONS****************************************************************
function amtd_GetBearerString_() {
//Call Amtd get access token using the refresh token - check validity of both access and refresh tokens.
// Access token lasts for 30 minutes, refresh token lasts for 90 days before having to require user to authenticate again
// curl -X POST --header "Content-Type: application/x-www-form-urlencoded" -d "grant_type=refresh_token&refresh_token=<refresh_token>&redirect_uri=https%3A%2F%2F127.0.0.1" "https://api.tdameritrade.com/v1/oauth2/token"

  var refresh_token = userProperties.getProperty("refresh_token");
  var refresh_time = userProperties.getProperty("refresh_time");
  var access_token = userProperties.getProperty("access_token");
  var access_time = userProperties.getProperty("access_time");
  var mynow = new Date();
  
  if ( (Date.parse(mynow) - Date.parse(access_time)) <29*60*1000 ) { //Access token is still not expired
    return "Bearer " + access_token;  
  } else if ( (Date.parse(mynow) - Date.parse(refresh_time)) >90*24*60*60*1000 ) {  //Refresh token expired
    //re-authenticate - amtd_showPane() ?
    return "Re-authentication needed!";    
  }
  
  var formData = {
    "grant_type" : "refresh_token",
    "refresh_token" : refresh_token,
    "client_id" : apikey
  }
  var options = {
    "method" : "post",
    "payload" : formData
  }
  var myurl="https://api.tdameritrade.com/v1/oauth2/token";
  var result=UrlFetchApp.fetch(myurl, options);

  //Parse JSON
  var contents = result.getContentText();
  var json = JSON.parse(contents);
  
  access_token = json["access_token"];
  userProperties.setProperty("access_token", access_token);
  userProperties.setProperty("access_time", access_time);
  
  if (json.hasOwnProperty("refresh_token")) {
    refresh_token = json["refresh_token"];
    userProperties.setProperty("refresh_token", refresh_token);
    userProperties.setProperty("refresh_time", refresh_time);
  }
  
  return "Bearer " + access_token;   
}

function amtd_GetTokens(s) {
//Receive the URI, strip out the code, and call Ameritrade to receive Bearer Token and Refresh Token
// Access token lasts for 30 minutes, refresh token lasts for 90 days before having to require user to authenticate again
// curl -X POST --header "Content-Type: application/x-www-form-urlencoded" -d "grant_type=refresh_token&refresh_token=<refresh_token>&redirect_uri=https%3A%2F%2F127.0.0.1" "https://api.tdameritrade.com/v1/oauth2/token"
   
//NB: Does not work when run in scripts. Should work if run in browser. 

  mycode = decodeURIComponent(s.substring(s.indexOf("code=")+5));
  
  var formData = {
    "grant_type" : "authorization_code",
    "access_type" : "offline",
    "code" : mycode,
    "client_id" : apikey,
    "redirect_uri" : "https://127.0.0.1"
  }
  var options = {
    "method" : "post",
    "payload" : formData
  }
  var myurl="https://api.tdameritrade.com/v1/oauth2/token";
  var result=UrlFetchApp.fetch(myurl, options);
  
//  {
//  "access_token" : "LsnQVY+UZ4JLZjZlUJbR/WcOsyiCCcjkv9QAVICUvkD8HuAIuUQsHDg1+rtsxsqGnBWKG+bMWyzJqmgvWn/GuYZCdl7RFBbqIuz1yDOFEsfUJSeiICfp9PaLU2h71gHnzXjuufxkDOFXtKy72SIdCrxllrL0Gx/mJiOnKkStRZLv7DKUy9WYq5E/njREFhdmFIDMfZ8ISgqLTNLO0SvDVaEmmmZ+jXZLMzqfKyrCqxjO56NEqCW2wGSkkJjSMEnD9NDOXo9FAQVDlWYpZaMiEA2IPXN+6YJ4zVDcFKpp7BCiJEGuJ/wB2OkcARtOt2qgXUfrm5qtxG6v2chXm2zf+nQcvVpFPTM5hUtCSnHlXmQUdlV595gpt8WRjuho2PXJGlANVd7A3nQRbtDF9OJc4onzsaRIbzCvcO/QGxzg1UcDiFwSnhJ3V8ComPnlKvNZ8FL65IRh8bdWDlKPjWziFVS7DiUUt1usuOvhbcZ1St8iP1F5SFtcdwkFo0Plzq8udKYLL6n+JOTOoIgczfRZvyUhDyPEwjEPBVlN3100MQuG4LYrgoVi/JHHvlZPz5Yp3HY5JTh+vCkJ4WbzO2X6Z9jkYmTnpm+5TO7ZMa7cAIoNosakkwtOy3PNpz+NdMqpTZk+y18DAxJEikARKtjTSIhOqMBhMh7G2qJD9a2IZYQGoGV8HJL63NZXLyFuvwJxmHGtolVkZVvBbfLIxwYWhHZGEICf/M44zGeHVLqxpkh9mLH5WGvznQEyEonPfhwH1JyEDx6ElcngR3ujB4B41htCj9d2oYbKulM9Vw4VbrrUyhqkZiK+JLZrqLpicG2sjDGT2FBsYXFasAF5q9R72Q5DoAmCnc+rkVfxCQWKsh2mb9lESmtuCsRHZolMhLP2Kd0ThwAltJk3lZlx8YQe3CtcGj3RO4XJ1799B1Lyy7PJKq6MmnHFT9+Wi4A5MV+j5M0us61ky5MbR5iRkpOmWaPZec9mmCxMkNxJXcUsP58cafZZBsZvHKB0uBdMBRWBa1FNvjCJkluTksBTr8SpgPrdifjpzo8yN/c52sJLrzCe4PpGZjbPyqEbB3OZcz+koZpRuXQUisOPgIDggQH+3Rc3RnZHPkyo212FD3x19z9sWBHDJACbC00B75E",
//  "refresh_token" : "fNuH0AR4vuxpBSwTrH7V527i4mQb6/Zxayra1XPlZdf3s6RxDabO5sWDA+Zi9i/lGociLF97r5HN6Xl22PqMplV8krhK9PpAuIK7FCFIIzscqtbF2ji4R8kPFzPvwmqFaE2QvxTC0azc63hs5bHaWISR0gsOHOtx/zHA20G+CFFwncqPVOHia9UT+WPu/zpRkAPfFPuRqKVj/7PlsDNqt7wZwtmZQN9Av4A6hzN56ZrKOovdnbaFftUYcX3gcbmMJOUharM1C9dmfP9n85LneGS2uivsiIBH3GVclIxGXNtsLnHoBHkiYCT4sUyv1PidynGAbi6jXcnApINwwf38PYx1LFi1a1YZrABdQGnJIs0uKFDtlpxNkuxonXsbHdvPeDOY82ouUPFkh2OydaKgz2i2r3LtMazgZFe5nj2/atfETz2XCd+kOUIVDoS100MQuG4LYrgoVi/JHHvlBbMb8YOaDxKBoZ3eX3ylTd+LEU3xDCvkBI+8rsbD9e4lTkjHwtEEOSK67Ijn85z9X0YBSaNiJ1M4QKgl3FCNRzwpzer4VhptkS91FHog3c/F03Vc1u3T8erfUcoF6Y0Uo7PwjhGT/+fp51JluoCVUvojtfjHGi6eXUpZ5iDxSaVTrK0lEPPA7bQXqxUXk5fP/m/0nT+jq+kXOwdHgIK8IWcWG1LZLpPuY8x4eoihXNC2FbHkKQBxkj3v0Q7J+bKj/orwXW/dMpHU3t1RH3bmPFJKZsaHei3Fh1W+eNmYyXw4v2oBzogda0nm09IPOXdR6So5NDeaTMqhKYstkw6rLf9tGzIxljiQwJHwWyX0tqkIa7AmrgCSDmXWbARUe2QhabmjJ92+sMi8wVZd9uhtSyPKi+lxrFUf1AneBGX3kwP6EOrpkXH221Uz4dE=212FD3x19z9sWBHDJACbC00B75E",
//  "scope" : "PlaceTrades AccountAccess MoveMoney",
//  "expires_in" : 1800,
//  "refresh_token_expires_in" : 7776000,
//  "token_type" : "Bearer"
//  }

  //Parse JSON
  var contents = result.getContentText();
  var json = JSON.parse(contents);
  
  access_token = json["access_token"];
  refresh_token = json["refresh_token"];
  
  userProperties.setProperty("access_token", access_token);
  userProperties.setProperty("access_time", new Date());
  userProperties.setProperty("refresh_token", refresh_token);
  userProperties.setProperty("refresh_time", new Date());
  
  //amtd_putTokens();  //For debugging 

  Logger.log("access_token : " + access_token + " refresh_token : " + refresh_token);
}

//*****************************UTILITY FUNCTIONS****************************************************************

function amtd_putTokens(tokensheet, rngAccessToken, rngRefreshToken) {
//put the access and refresh tokens and their times from userProperties in the spreadsheet

//  var tokensheet = "Example";    //For Ameritrade API w Google Scripts
//  var rngAccessToken = "D9";
//  var rngRefreshToken = "D10";
  
//  var tokensheet = "Template";    //For StockPortfolio
//  var rngAccessToken = "D19";
//  var rngRefreshToken = "D20";
  
  var currentssht = SpreadsheetApp.getActive();
  var sourcesht = currentssht.getSheetByName(tokensheet);
  
  var access_token = userProperties.getProperty("access_token");
  var access_time = userProperties.getProperty("access_time");
  var refresh_token = userProperties.getProperty("refresh_token");
  var refresh_time = userProperties.getProperty("refresh_time");
  
  sourcesht.getRange(rngAccessToken).setValue(access_token);
  sourcesht.getRange(rngAccessToken).offset(0, -1).setValue(access_time);
  sourcesht.getRange(rngRefreshToken).setValue(refresh_token);
  sourcesht.getRange(rngRefreshToken).offset(0, -1).setValue(refresh_time);
}

function epochdate_(thisdate) {
  //Return thisdate in msec in string format
//  thisdate = "1/27/2020";
  var dDate = Date.parse(thisdate).toString();
//  var dDate = (thisdate).toString();
  return dDate;     //.toString();
//  slog(dDate);
//  var datestr = new Date(dDate).toDateString();
//  slog(datestr);
}
