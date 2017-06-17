/**
 * @OnlyCurrentDoc
 */

function onOpen() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var entries = [{ name : "Update", functionName : "updateData" },
                 { name : "Clear", functionName : "clearData" }]
  sheet.addMenu("Settings", entries);
}

function updateData(){
  SpreadsheetApp.getActiveSpreadsheet().getRange("S2").setValue(new Date())
}

function clearData(){
  const d = new Date()
  SpreadsheetApp.getActiveSpreadsheet().getRange("A2").setValue("")
  SpreadsheetApp.getActiveSpreadsheet().getRange("H2").setValue(2)
  SpreadsheetApp.getActiveSpreadsheet().getRange("G2").setValue(new Date(d.getYear(), d.getMonth(), d.getDate()+7, 15, 30, 0))
  SpreadsheetApp.getActiveSpreadsheet().getRange("B6:G15").setValue("")
  SpreadsheetApp.getActiveSpreadsheet().getRange("S2").setValue(d)
}


/**
Get list of options expiry dates
@customFunction
*/
function _getExpiries(symbol){
  const e = JSON.parse(UrlFetchApp.fetch("https://www.nseindia.com/live_market/dynaContent/live_watch/get_quote/ajaxFOGetQuoteDataTest.jsp?i="+(symbol=="NIFTY"||symbol=="BANKNIFTY"?"OPTIDX":"OPTSTK")+"&u="+symbol).getContentText())["expiries"]
  return e.map(function(x){
    return new Date(+x.slice(5,9),["JAN","FEB","MAR","APR","MAY","JUN","JUL","AUG","SEP","OCT","NOV","DEC"].indexOf(x.slice(2,5)),+x.slice(0,2), 15, 30, 0)
  }).sort(function(d1,d2){return d1>d2?1:-1})
}

/**
Get list of avilable strikes
@customFunction
*/
function _getStrikes(symbol){
  const e = _getExpiries(symbol)[0]
  const s = JSON.parse(UrlFetchApp.fetch("http://www.nseindia.com/live_market/dynaContent/live_watch/get_quote/ajaxFOGetQuoteDataTest.jsp?i="+(symbol=="NIFTY"||symbol=="BANKNIFTY"?"OPTIDX":"OPTSTK")+"&u="+symbol+"&e="+(e.getDate().toString()+["JAN","FEB","MAR","APR","MAY","JUN","JUL","AUG","SEP","OCT","NOV","DEC"][e.getMonth()]+e.getYear().toString())+"&o=CE&k=CE").getContentText())["strikePrices"]
  return s.map(Number).sort(function(n1,n2){return n1>n2?1:-1})
}

/**
Get futures price and lot size for a symbol-expiry pair
@customFunction
*/
function _getFutureAndLotSize(symbol, expiry, _){
  const url = "https://www.nseindia.com/live_market/dynaContent/live_watch/get_quote/ajaxFOGetQuoteJSON.jsp?underlying="+symbol+"&instrument="+(symbol=="NIFTY"||symbol=="BANKNIFTY"?"FUTIDX":"FUTSTK")+"&expiry="+(expiry.getDate().toString()+["JAN","FEB","MAR","APR","MAY","JUN","JUL","AUG","SEP","OCT","NOV","DEC"][expiry.getMonth()]+expiry.getYear().toString())+"&type="+"&strike="
  const f = JSON.parse(UrlFetchApp.fetch(url, {"headers":{"Referer":"https://www.nseindia.com","User-Agent":"Mozilla/5.0","Accept":"/"}}).getContentText()).data[0]
  return [f.lastPrice.replace(/,/g, ""), f.marketLot].map(Number)
}

/**
Get options quotes [LTP, pClose, Bid, Ask, OI, ΔOI, IV]
@customFunction
*/
function _getOptionQuotes(symbol, expiry, strike, call_or_put, _){
  const url = "https://www.nseindia.com/live_market/dynaContent/live_watch/get_quote/ajaxFOGetQuoteJSON.jsp?underlying="+symbol+"&instrument="+(symbol=="NIFTY"||symbol=="BANKNIFTY"?"OPTIDX":"OPTSTK")+"&expiry="+(expiry.getDate().toString()+["JAN","FEB","MAR","APR","MAY","JUN","JUL","AUG","SEP","OCT","NOV","DEC"][expiry.getMonth()]+expiry.getYear().toString())+"&type="+call_or_put+"&strike="+parseFloat(strike).toFixed(2)
  const q = JSON.parse(UrlFetchApp.fetch(url, {"headers":{"Referer":"https://www.nseindia.com","User-Agent":"Mozilla/5.0","Accept":"/"}}).getContentText()).data[0]
  return [q.lastPrice, q.prevClose, q.buyPrice1, q.sellPrice1, q.openInterest.replace(/,/g, ""), q.changeinOpenInterest.replace(/,/g, ""), q.impliedVolatility].map(Number)
}

/**
Calculate Black-Scholes price
@customFunction
*/
function _getBS(futurePrice, strike, time_to_expiry, volatility, call_or_put){
  const r = 0
  const d1 = (Math.log(futurePrice / strike) + (r + volatility * volatility / 2) * time_to_expiry) / (volatility * Math.sqrt(time_to_expiry))
  const d2 = d1 - volatility * Math.sqrt(time_to_expiry);
  return call_or_put === "CE" ? futurePrice * CND(d1)-strike * Math.exp(-r * time_to_expiry) * CND(d2) : strike * Math.exp(-r * time_to_expiry) * CND(-d2) - futurePrice * CND(-d1)
}

/**
Get implied volatility of an option
@customFunction
*/
function _getIV(optionPrice, futurePrice, strike, time_to_expiry, call_or_put)
{
  var iv = 0.1, low = 0, high = Infinity
  for(var i = 0; i < 100; i++){
    var bsPrice = _getBS(futurePrice, strike, time_to_expiry, iv, call_or_put)
    if(iv*100 === Math.floor(bsPrice*100)) break
    else if (bsPrice > optionPrice){
      high = iv
      iv = (iv-low)/2 + low
    } else {
      low = iv
      iv = (high-iv)/2 + iv
      if(!isFinite(iv)) iv = low * 2;
    }
  }
  return iv
}

/**
Get option greeks [Delta, Theta, Gamma, Vega]
@customFunction
*/
function _getGreeks(futurePrice, strike, time_to_expiry, volatility, call_or_put){
  return [bsDelta(futurePrice, strike, time_to_expiry, volatility, call_or_put),
          bsTheta(futurePrice, strike, time_to_expiry, volatility, call_or_put),
          bsVega(futurePrice, strike, time_to_expiry, volatility),
          bsGamma(futurePrice, strike, time_to_expiry, volatility)]
}


/**
Generate price range or payoff table
@customFunction
*/
function _payoffRange(payoff, eval, future, iv, range){
  payoff = new Date (payoff)
  eval = new Date(eval)
  const timediff = (payoff-eval)/(1000*60*60*24)
  const sd = iv/Math.sqrt((365/timediff))
  const min = future*(1-range*sd)
  const max = future*(1+range*sd)
  var prices = []
  for (i=0; i<50; i++){
    prices[i] = (min+i*(max-min)/50)
  }
  return prices
}


/**
Generate payoff series
@customFunction
*/
function _getPayoff(payoffRange, payoff, eval, expiry, strike, call_or_put, iv){
  expiry = new Date(expiry)
  payoff = new Date(payoff)
  eval = new Date(eval)
  const t = Math.max((expiry-payoff)/(1000*60*60*24), 0)/365
  var prices = []
  for (i=0; i<50; i++){
    prices.push(_getBS(payoffRange[i], strike, t, iv, call_or_put))
  }
  return prices
}


/**
Get margin requirement
@customFunction
*/
function _getMargin(symbol, lotsize, lots, expiries, strikes, call_or_put, _){
  if ([symbol].length != 1) return "Invalid Symbol"
  else if ([lotsize].length != 1 || typeof lotsize != "number" ) return "Invalid Lot Size"
  lots = lots.filter(function(x) { return /\S/.test(x) })
  var data = []
  for (var i=0; i<lots.length; i++){
    date = new Date(expiries[i])
    data.push({"Id" : symbol+date.getYear().toString().slice(2,4)+["JAN","FEB","MAR","APR","MAY","JUN","JUL","AUG","SEP","OCT","NOV","DEC"][date.getMonth()]+strikes[i]+call_or_put[i]+"|NFO|",
               "qty" : Number(lots[i])})
  }
  
  var grouped_data = []
  data.reduce(function (res, value){
    if (!res[value.Id]){
      res[value.Id] = {
        qty: 0,
        Id: value.Id
      }
      grouped_data.push(res[value.Id])
    }
    res[value.Id].qty += value.qty
    return res}, {})

  var q = '{"InputStr":"'
  for (var i=0; i<grouped_data.length; i++){
    i>0 ? q+="%23" : ""
    q+= grouped_data[i].Id + lotsize*grouped_data[i].qty + "|0|0"
  }
  q+= '"}'

  const url = "http://www.karvyonline.com/serviceapis/ServiceHelper.asmx/GetSpanInterop"
  const options = { 'method' : 'post',
                  'payload': q,
                  'headers' : { 'Content-Type': 'application/json; charset=UTF-8' }
                };
  const response = UrlFetchApp.fetch(url, options);
  const f = JSON.parse(JSON.parse(response.getContentText())["d"])
  var margin = (+f[0].totalrequirement) + (+f[0].exposuremargin)
  
  return margin
}

 


//------- SUPPORT FUNCTIONS -------


function CND(x){
  if (x < 0) return 1-CND(-x)
  k = 1/(1+ 0.2316419 * x)
  return (1 - Math.exp(-x*x/2) / Math.sqrt(2*Math.PI) * k * (.31938153 + k * (-.356563782 + k * (1.781477937 + k * (-1.821255978 + k * 1.330274429)))))
}

function ND(x){
  return 1.0/Math.sqrt(2*Math.PI) * Math.exp(-x * x / 2)
}

function bsDelta(futurePrice, strike, time_to_expiry, volatility, call_or_put){
  const r = 0
  const d1 = (Math.log(futurePrice / strike) + (r + volatility * volatility / 2) * time_to_expiry) / (volatility * Math.sqrt(time_to_expiry))
  return call_or_put === "CE" ? Math.exp((r) * time_to_expiry) * CND(d1) : Math.exp((r) * time_to_expiry) * (CND(d1)-1)
}

function bsGamma(futurePrice, strike, time_to_expiry, volatility){
  const r = 0
  const d1 = (Math.log(futurePrice / strike) + (r + volatility * volatility / 2) * time_to_expiry) / (volatility * Math.sqrt(time_to_expiry))
  return ND(d1) * Math.exp((r) * time_to_expiry) / (futurePrice * volatility * Math.sqrt(time_to_expiry))
}

function bsVega(futurePrice, strike, time_to_expiry, volatility){
  const r = 0
  const d1 = (Math.log(futurePrice / strike) + (r + volatility * volatility / 2) * time_to_expiry) / (volatility * Math.sqrt(time_to_expiry))
  return (futurePrice * Math.exp((r) * time_to_expiry)  * ND(d1) * Math.sqrt(time_to_expiry))/100
}

function bsTheta(futurePrice, strike, time_to_expiry, volatility, call_or_put){
  const r = 0
  const d1 = (Math.log(futurePrice / strike) + (r + volatility * volatility / 2) * time_to_expiry) / (volatility * Math.sqrt(time_to_expiry))
  const d2 = d1 - volatility * Math.sqrt(time_to_expiry)
  const c = -futurePrice * Math.exp((r) * time_to_expiry)  * ND(d1) * volatility/(2.0 * Math.sqrt(time_to_expiry)) - (r) * futurePrice * Math.exp((r) * time_to_expiry) * CND(d1) -r * strike * Math.exp(-r * time_to_expiry)  * CND(d2)
  const p = -futurePrice * Math.exp((r) * time_to_expiry)  * ND(d1) * volatility/(2.0 * Math.sqrt(time_to_expiry)) + (r) * futurePrice * Math.exp((r) * time_to_expiry) * CND(-d1) +r * strike * Math.exp(-r * time_to_expiry)  * CND(-d2)
  return call_or_put === "CE" ? c/365 : p/365 
}