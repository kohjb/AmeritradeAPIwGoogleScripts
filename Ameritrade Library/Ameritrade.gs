
var apikey = "YourAPIKeyHere....";
var userProperties = PropertiesService.getUserProperties();



//****************************** MAIN FUNCTIONS *********************************************

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
  
  Price = stock["closePrice"];
  // Other possible options are bidPrice, ask Price, lastPrice, openPrice, highPrice, lowPrice, etc. See Schema.
  
  return Price;
}

/**
 * Call Ameritrade API to get the closing prices of one or more stockSymbols.
 *
 * @param {string} the stock's symbol - comma seperated list of symbols, e.g. "AAPL,/EQ"
 * @return {string} the closing price - comma seperated list of prices, e.g. "167.30,4369.50"
 */
 function amtd_GetQuotes(stockSymbols) {

  var authorization = amtd_GetBearerString_();
  var options = {
    "method" : "GET",
    "headers" :  {"Authorization" : authorization},
    "apikey" : apikey,
  }
  var foptions = "?symbol="+stockSymbols;
  var myurl="https://api.tdameritrade.com/v1/marketdata/quotes"+foptions;
  var result=UrlFetchApp.fetch(myurl, options);

  //Parse JSON
  var contents = result.getContentText();  //If need to debug what the returned result is
  var json = JSON.parse(contents);
  //return JSON.stringify(json);  //If need to return json as a string
  //var stock = json[stockSymbol];

  var Prices = ""
  for (let x in json) {  //Go through each symbol in json, get price depending on assetType
    if (Prices != "") Prices += ","  //If not the first price, add comma seperator
    switch (json[x]["assetType"]) {
      case "EQUITY":
        Prices += json[x]["lastPrice"]
        break
      case "FUTURE":
        Prices += json[x]["lastPriceInDouble"]
        break
    }
  }
  return Prices
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
//Call Ameritrade API and return the Option's price for expDateStart only

  //For debugging
//  var stockSymbol = "JNJ";
//  var contractType = "CALL";
//  var strike = 150;
//  var expDate = "2020-05-15";
//  var expDateEnd = "2020-02-21";
//  var priceType = "bid";
//  expDateStart = "2020-05-15";
  expDateStart = Utilities.formatDate(new Date(expDate), "GMT+8", "yyyy-MM-dd").toString();
  expDateEnd = expDateStart; //"2020-07-17"; 
  var authorization = amtd_GetBearerString_();
  var options = {
    "method" : "GET",
    "headers" :  {"Authorization" : authorization},
    "apikey" : apikey
  }
  var foptions = "?symbol="+stockSymbol+"&contractType="+contractType+"&strike="+strike+"&fromDate="+expDateStart+"&toDate="+expDateEnd;
  
  var myurl="https://api.tdameritrade.com/v1/marketdata/chains"+foptions;

  var result=UrlFetchApp.fetch(myurl, options);
  
  //Parse JSON
  var contents = result.getContentText();
  var json = JSON.parse(contents);
  Logger.log(json);
  //Sample JSON2:
    //[20-05-04 14:15:57:923 HKT] {interestRate=0.11, symbol=JNJ, isDelayed=true, underlyingPrice=148.165, putExpDateMap={}, daysToExpiration=0, underlying=null, isIndex=false, 
    //volatility=29, callExpDateMap={2020-05-15:11={150=[{totalVolume=269, symbol=JNJ_051520C150, openInterest=23718, optionDeliverablesList=null, delta=0.417, 
    //description=JNJ May 15 2020 150 Call, openPrice=0, volatility=30.462, timeValue=2.22, theta=-0.129, lastTradingDay=1.5895872E12, lowPrice=1.78, highPrice=2.52, 
    //askSize=19, theoreticalVolatility=29, expirationDate=1.5895728E12, markPercentChange=0, netChange=-0.17, settlementType= , isIndexOption=null, percentChange=-7.31, 
    //expirationType=R, last=2.22, mini=false, bidSize=40, multiplier=100, daysToExpiration=11, inTheMoney=false, tradeTimeInLong=1.588363055412E12, tradeDate=null, putCall=CALL, 
    //quoteTimeInLong=1.588363199909E12, markChange=0, lastSize=0, nonStandard=false, ask=2.56, rho=0.019, exchangeName=OPR, deliverableNote=, closePrice=2.39, bid=2.23, 
    //bidAskSize=40X19, theoreticalOptionValue=2.395, mark=2.4, gamma=0.048, vega=0.105, strikePrice=150}]}}, interval=0, strategy=SINGLE, numberOfContracts=1, status=SUCCESS}
  
  var status = json["status"];
  if (status == "SUCCESS") {
    mapname = contractType.toLowerCase() + "ExpDateMap";
    
    putcallmap = json[mapname];
    
    for(var k1 in putcallmap) if (k1.indexOf(expDateStart)>-1) {
      datemap = putcallmap[k1];
      for(var k2 in datemap) if (k2.indexOf(strike.toString())>-1) {
        pricemap = datemap[k2];
        price = pricemap[0];
        optprice = price[priceType];
        return optprice;
      } else {
        return "Strike not found";
      }
    } else {
      return "Putcallmap with date note found";
    }
  } else {
    return "FAIL";
  }
}

/**
 * Call Ameritrade API and return an object containing the Option's chain of information.
 *
 * @param {string} the stock's symbol
 * @param {string} CALL or PUT contract
 * @param {number} the strike price of the contract
 * @param {date} the expiry date of the contract
 * @param {string} bid or ask price
 * @return {number} the option's price
 */
function amtd_GetOptionPriceChain(stockSymbol, contractType, strike, expDateStart, expDateEnd) {
//Call Ameritrade API and return an object containing the Option's chain of information.

  //For debugging
//  var stockSymbol = "GOOG";
//  var contractType = "PUT";
//  var strike = 1450;
//  var expDateStart = "2020-03-21";
//  var expDateEnd = "2021-01-23";

  expDateStart = Utilities.formatDate(expDateStart, "GMT+8", "yyyy-MM-dd").toString();
  expDateEnd = Utilities.formatDate(expDateEnd, "GMT+8", "yyyy-MM-dd").toString();
  var authorization = amtd_GetBearerString_();
  var options = {
    "method" : "GET",
    "headers" :  {"Authorization" : authorization},
    "apikey" : apikey
  }
  var foptions = "?symbol="+stockSymbol+"&contractType="+contractType+"&strike="+strike+"&fromDate="+expDateStart+"&toDate="+expDateEnd;
  
  var myurl="https://api.tdameritrade.com/v1/marketdata/chains"+foptions;

  var result=UrlFetchApp.fetch(myurl, options);
  
  return result;
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
//  slog("GPH.options = " + options);
//  return(foptions);
  
  //Create url and fetch it
  var myurl="https://api.tdameritrade.com/v1/marketdata/"+stockSymbol+"/pricehistory"+foptions;
  var result=UrlFetchApp.fetch(myurl, options);  
  //slog(result);   

  //Parse JSON
  var contents = result.getContentText();
  var json = JSON.parse(contents);
//  return(JSON.stringify(json));
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
  
  //Continue here if Access token has expired but Refresh Token has not expired.
  //Check if Refresh Token is about to expire, and if so, get a new Refresh Token as well.

  if ( (Date.parse(mynow) - Date.parse(refresh_time)) >80 *24*60*60*1000 ) {  //Refresh token greater than 80 days old
    //Get new Refresh Token + Access Token
    var formData = {
    "grant_type" : "refresh_token",
    "refresh_token" : refresh_token,
    "access_type" : "offline",
    "client_id" : apikey
    }
  } else {
    //Get new Access Token only
    var formData = {
    "grant_type" : "refresh_token",
    "refresh_token" : refresh_token,
    "client_id" : apikey
    }
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
  userProperties.setProperty("access_time", new Date());
  
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

  //https://127.0.0.1/?code=xm28xeurnCu2Hc6ME%2FyU8xxLuPqLu1toon0kV6SZoURCqaW0Nf3Mm6h45ixRRYOl2W8kDHAwEG4kEJEHV8iXwYcZYxl8aqBD4EIwW%2FBVuLd%2BjAWlogJQ292aqFFBdz8iotEsWVOVbOtJRytGGN8IdnrSQhpxVYIvCffmN%2BfJNcoyg%2Bzx6lh6vhXiSI6fyjLhPjDazux4tFE9%2FigXbyNXJykiZCiRXayUNGE1AsOez5N%2Bi3HEsWZYy5eSlHTqtk7NbDtTuw18DVkYKBFGNg%2FcuW9gKE%2BtZCp9Gd6jGpgBbkucd5L7stif%2BeFBM8DTKimqf0gunbcv37c91JgykVXa5%2Fr0Wc3BTtBoWpINGKU4%2BN%2FrEobya2cCj%2BlUYR1yc9iG%2FywN6AvA%2BIZFRyLZotdg0Mcaxt62VIR9yA41umGHKv2SVp65e65KKrBjtKx100MQuG4LYrgoVi%2FJHHvllG0yTDs31857vxtpjDIy4mmoEuZDf%2FT9pUCwNEA8DOH9DoadKWImK74ivtJopZzQkuplSURrDtxBpMRK5lLZyCsb10K5aIz9gx9fOqvxXMK3QHr04ga%2F19%2FAfp1DZvhxqBt3YrxyDpS5xO6yWS7RDWQPnCJXMXJz%2FdsctugNaWm2See3bavUruhlHpnVrOk%2FhoQsrEr8xpLPpisiJEdMiA1xo8HIvqMJ8wOc7ClrTDaH67ZAfRJZCmnUqyk8b20%2BcY%2BabgH%2BPnVPNrQ6wlpLYVJbuNA%2FqBEtpBFSdd0LvNeJKoS5xA8ynk1x7oStJSVOhzZJwyq7FemBbBSfLhNJa1JYRHPtsZAyuSu75oqjnHMlXvfSBlegzl3EisJm3STrSNScpnvb%2Bg%2FU6pADLEqZRZ1uzlyWPpI18HHwInejGXH3g%2FGlCDYF4EVI2w4%3D212FD3x19z9sWBHDJACbC00B75E
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

/**
 * Get the Access Token from userProperties
 *
 * @return {string}
 */
function amtd_getAccessToken() {
//Get the Access Token from userProperties
  return userProperties.getProperty("access_token");
}

/**
 * Get the time the Access Token was requested from userProperties
 *
 * @return {date}
 */
function amtd_getAccessTokenTime() {
//Get the Access Token Time from userProperties
  return userProperties.getProperty("access_time");
}

/**
 * Get the Refresh Token from userProperties
 *
 * @return {string}
 */
function amtd_getRefreshToken() {
//Get the Refresh Token from userProperties
  return userProperties.getProperty("refresh_token");
}

/**
 * Get the time the Refresh Token was requested from userProperties
 *
 * @return {date}
 */
function amtd_getRefreshTokenTime() {
//Get the Refresh Token Time from userProperties
  return userProperties.getProperty("refresh_time");
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

function expdate_(d) {
//Return a date d in the format that Option date requires.

  //d=new Date();
  var fd = Utilities.formatDate(d, "GMT+8", "yyyy-MM-dd").toString();
  Logger.log(fd);
  return fd;
}










