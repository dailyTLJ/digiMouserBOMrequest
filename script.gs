
const MOUSER_API_KEY = "***************************************";
const MOUSER_KEYWORD_SEARCH_UL = `https://api.mouser.com/api/v1/search/keyword?apiKey=${MOUSER_API_KEY}`;

const DIGIKEY_CLIENT_ID = '***************************************';
const DIGIKEY_CLIENT_SECRET = '***************';
const DIGIKEY_AUTH_URL_V4 = 'https://api.digikey.com/v1/oauth2/authorize';
const DIGIKEY_TOKEN_URL_V4 = 'https://api.digikey.com/v1/oauth2/token';
const DIGIKEY_PRODUCT_SEARCH_URL_V4 = 'https://api.digikey.com/products/v4/search/keyword';



var mouserCol = 2;
var digikeyCol = 6;
var checkboxCol = 11;
var logCol = 12;


// curl -X POST "https://api.mouser.com/api/v1/search/keyword?apiKey=xxxxxxxxxxxxxxxxx" -H "accept: application/json" -H "Content-Type: application/json" -d "{ \"SearchByKeywordRequest\": { \"keyword\": \"CA11976_LAURA-M\", \"records\": 0, \"startingRecord\": 0, \"searchOptions\": \"string\", \"searchWithYourSignUpLanguage\": \"string\" }}"

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Custom Menu')
      .addItem('updateAll', 'updateAll')
      .addToUi();
}


function test() {
  updateRow(3);
}


function updateAll() {

  var sheet = SpreadsheetApp.getActiveSheet();
  var rows = sheet.getDataRange().getValues();

  rows.forEach(function(row, index) {
    if (index < 2) return;
    updateRow(index+1);

  });

  var currentDate = new Date();
  var formattedDate = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm");
  sheet.getRange(1, 1).setValue("last updated: "+formattedDate);

  log("");  // to clear the log field
  
}




function updateRow(rowID) {

  var sheet = SpreadsheetApp.getActiveSheet();
  var row = sheet.getRange(rowID,1,1,1).getValues();

  sheet.getRange(rowID,logCol).setValue(`updating`);


  const partNumber = row[0][0];
  Logger.log("row "+rowID+": "+partNumber);

  if (!partNumber) return;

  const partMouser = callMouserAPI(partNumber);
  // Logger.log(partMouser);
  
  if(partMouser) {
    sheet.getRange(rowID,logCol).setValue(`add Mouser`);

    // Extract relevant fields
    const availability = partMouser.Availability || 'N/A';
    const dataSheetUrl = partMouser.DataSheetUrl || 'N/A';
    const imagePath = partMouser.ImagePath || 'N/A';
    const priceForQty1 = partMouser.PriceBreaks?.find(pb => pb.Quantity === 1)?.Price || 'N/A';
    const availabilityInStock = partMouser.AvailabilityInStock || 'N/A';
    const factoryStock = partMouser.FactoryStock || 'N/A';
    const leadTime = partMouser.LeadTime || 'N/A';
    const productDetailUrl = partMouser.ProductDetailUrl || 'N/A';
    
    sheet.getRange(rowID,logCol).setValue(`writeMouser`);
    sheet.getRange(rowID, mouserCol, 1, 4).setValues([[
      priceForQty1, availabilityInStock, leadTime, `=hyperlink("${productDetailUrl}","mouser.ca")`
    ]]);
    

  }

  const partDigikey = callDigikeyAPI(partNumber);
  // Logger.log(partDigikey);

  if(partDigikey) {

    // sheet.getRange(rowID,logCol).setValue(`add Digikey`);

    const quantityAvailable = partDigikey.QuantityAvailable || 0.0;
    const manufacturerLeadWeeks = `${partDigikey.ManufacturerLeadWeeks} weeks` || "N/A";
    const unitPrice = partDigikey.UnitPrice || "N/A";
    const productUrl = partDigikey.ProductUrl || "N/A";
    productUrl.toString().replace("digikey.com", "digikey.ca");
    let updatedUrl = productUrl.replace(/digikey\.com/, "digikey.ca");
    Logger.log(updatedUrl);
    sheet.getRange(rowID,logCol).setValue(`writeDigikey`);
    sheet.getRange(rowID, digikeyCol, 1, 4).setValues([[
      unitPrice, quantityAvailable, manufacturerLeadWeeks, `=hyperlink("${updatedUrl}","digikey.ca")`
    ]]);
    

  }


  sheet.getRange(rowID,logCol).setValue(``);
  
}



function log(line) {
  var sheet = SpreadsheetApp.getActiveSheet();
  Logger.log(line);
  sheet.getRange(1,logCol).setValue(line);

  var logSheet = SpreadsheetApp.getActive().getSheetByName('log')
  var content = logSheet.getRange(1,1).getValue();
  logSheet.getRange(1,1).setValue(content + line);
}


function callMouserAPI(partNumber) {

  log(`Mouser API request for part ${partNumber}`);

  // Prepare the POST request payload
  const payload = {
    SearchByKeywordRequest: {
      keyword: partNumber,
      records: 1,
      startingRecord: 0,
      searchOptions: "",
      searchWithYourSignUpLanguage: ""
    }
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload)
  };
  
  try {
    // Fetch data from the Mouser API
    const response = UrlFetchApp.fetch(MOUSER_KEYWORD_SEARCH_UL, options);
    const jsonResponse = JSON.parse(response.getContentText());
    return jsonResponse?.SearchResults?.Parts?.[0] || null;

  } catch (error) {
    log(`Error fetching data for ${partNumber}: ${error.message}`);
  }

}



function callDigikeyAPI(partNumber) {

  log(`Digikey API request for part ${partNumber}`);

  if (!DIGIKEY_CLIENT_ID || !DIGIKEY_CLIENT_SECRET) {
    log("Missing Digi-Key client ID or secret.");
    return;
  }

  // Step 1: Get the OAuth token
  const oauthToken = getOAuthToken();
  if (!oauthToken) {
    log("Failed to retrieve OAuth token.");
    return;
  }
  
  try {
    const searchResult = productSearch(oauthToken,partNumber);
    const jsonResponse = JSON.parse(searchResult.getContentText());
    
    const products = jsonResponse.ExactMatches;
    
    const product = products[0]; // Extract the first product
    Logger.log(product);
    return product;

  } catch (error) {
    log(`Error fetching data for ${partNumber}: ${error.message}`);
  }



}

//Get OAuth 2.0 access token for Digi-Key API.
function getOAuthToken() {
  const options = {
    method: 'post',
    payload: {
      client_id: DIGIKEY_CLIENT_ID,
      client_secret: DIGIKEY_CLIENT_SECRET,
      grant_type: 'client_credentials'
    }
  };

  const response = UrlFetchApp.fetch(DIGIKEY_TOKEN_URL_V4, options);
  const responseData = JSON.parse(response.getContentText());
  
  return responseData.access_token ? responseData : null;
}

// Perform a product search using the Digi-Key API.
function productSearch(token, keyword) {

  const payload = {
    Keywords: keyword
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      "X-DIGIKEY-Client-Id": DIGIKEY_CLIENT_ID,
      "X-DIGIKEY-Locale-Language": "en",
      "X-DIGIKEY-Locale-Currency": "CAD",
      "X-DIGIKEY-Locale-Site": "CA",
      "Authorization": "Bearer " + token.access_token
    },
    payload: JSON.stringify(payload)
  };

  const response = UrlFetchApp.fetch(DIGIKEY_PRODUCT_SEARCH_URL_V4, options);
  return response;
}






// function onEdit(e) {

//   var sheet = SpreadsheetApp.getActiveSheet();
//   var rows = sheet.getDataRange().getValues();

//   rows.forEach(function(row, index) {
//     if (index < 2) return;
//     const checkBox = row[checkboxCol-1];
//     if(checkBox === true) {
//       updateRow(index+1);
//       log('done');
//       sheet.getRange(index+1,logCol).setValue('Done.');
//       sheet.getRange(index+1,checkboxCol).setValue(false);;
//     }

//   });

  
// }