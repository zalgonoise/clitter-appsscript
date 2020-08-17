function initSheet() {

    // Open the associated Spreadsheet on the first sheet

  var file = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = file.getSheets()[0];
  

    // Defines headers for the table if non-existent

  if (sheet.getRange("A1").getValue() != "ID") {
    sheet.getRange("A1").setValue("ID")  
  }
  
  if (sheet.getRange("B1").getValue() != "Message") {
    sheet.getRange("B1").setValue("Message")  
  }
  
  if (sheet.getRange("C1").getValue() != "Date") {
    sheet.getRange("C1").setValue("Date")  
  }
  
    if (sheet.getRange("D1").getValue() != "Retweets") {
    sheet.getRange("D1").setValue("Retweets")  
  }
  
    if (sheet.getRange("E1").getValue() != "Favorited") {
    sheet.getRange("E1").setValue("Favorited")  
  }
  console.log("Initialized spreadsheet's headers")

}


  function getLatest() {

    // Getting the latest value present in the sheet 
    // by looking through all the Tweet ID cells 
    // and storing the last value

  var range = "A2:A50000"
  
    // Open the associated Spreadsheet on the first sheet

  var file = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = file.getSheets()[0];

  var cells = sheet.getRange(range).getValues();
  
    // Loops through each cell and stores its ID 
    // while the cell isn't empty, also storing the
    // empty cell number

   for (var i = 0 ; i < cells.length ; i++) {
    if (cells[i][0] === "" && !blank) {
      var blank = true
      var blankRow = (i+2)
      break
    } else {
      var blank = false
      var lastValue = cells[i][0]
    }
   }

    // In case there are no Tweets, all values are fetched
    
    if (!lastValue) {
      var lastValue = 0
      console.log("No values found. Fetching all that is reachable")
    }
    
    // Returns both the blank row number 
    // and the last Tweet ID

  return [blankRow, lastValue]
  }



function fetchElog() {  

    // Open the associated Spreadsheet on the first sheet

  var file = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = file.getSheets()[0];

    // First checks whether the Sheet is new and initialize it
  
  if (sheet.getRange("A1").getValue() === "") {
    console.log("Sheet seems blank, initializing.")
    initSheet()
    
  } else {
    console.log("Sheet check: OK")
  }
  
    // Fetches latest values and splits which is the next empty 
    // row as well as which is the last Tweet ID
  
  var getLatestValues = getLatest()
  var nextRow = getLatestValues[0]
  var latestID = getLatestValues[1] 

  console.log("nextRow: " + nextRow)
  console.log("latestID: " + latestID)

  
    // ACTION REQUIRED: Generate your Twitter API key [:1]
    // then replace the values in the variables below
    // [:1] https://developer.twitter.com/en/docs/authentication/oauth-1-0a/obtaining-user-access-tokens


  var twitterAPIKey = "YOUR_TWITTER_API_KEY";
  var twitterScrKey = "YOUR_TWITTER_API_SECRET"

    // Authenticates the user via OAuth 2.0 
  
  var tokenUrl = "https://api.twitter.com/oauth2/token";
  var tokenCredential = Utilities.base64EncodeWebSafe(
    twitterAPIKey + ":" + twitterScrKey
  );

  var tokenOptions = {
    headers : {
      Authorization: "Basic " + tokenCredential,
      "Content-Type": "application/x-www-form-urlencoded;charset=UTF-8"
    },
    method: "post",
    payload: "grant_type=client_credentials"
  };

  var responseToken = UrlFetchApp.fetch(tokenUrl, tokenOptions);
  var parsedToken = JSON.parse(responseToken);
  var token = parsedToken.access_token;

    // ACTION REQUIRED: Replace the variable below with the 
    // Twitter hanger you wish to collect data from 
  
  var twitterName = "elonmusk"

    // Builds the API request URL
  
  var apiUrl = "https://api.twitter.com/1.1/statuses/user_timeline.json?screen_name=" + twitterName + "&count=200&include_rts=true";

    // Appends the latest Tweet ID if existing, to trim results

  if (latestID != 0) {
   var apiUrl = apiUrl + "&since_id=" + latestID; 
  }
  
    // Fetches the requested data using the generated refresh token

  var apiOptions = {
    headers : {
      Authorization: 'Bearer ' + token
    },
    "method" : "get"
  };
  
  var responseApi = UrlFetchApp.fetch(apiUrl, apiOptions);

    // Initializes the result variable

  var result = "";

    // If the response is 200 (status code OK), then parse the results

  if (responseApi.getResponseCode() == 200) {
    var tweets = JSON.parse(responseApi.getContentText());

    if (tweets) {
      console.log("Number of found entries: " + tweets.length)

    // Loops through each fetched tweet, formats to a valid date
    // and sets the variables to be pushed into the spreadsheet

      for (var i = (tweets.length -1) ; i > -1; i--) {
        var tweetMsg = tweets[i].text;
        var tweetDate = new Date(tweets[i].created_at.toString().replace(/^\w+ (\w+) (\d+) ([\d:]+) \+0000 (\d+)$/, "$1 $2 $4 $3 UTC"));
        var tweetID = tweets[i].id;
        var tweetRTCount = tweets[i].retweet_count
        var tweetFavCount = tweets[i].favorite_count
        
        var newRow = nextRow;
        
    // Second validation for the first entry (to avoid duplicates)
    // performs a date check against the tweet and the existing content
    // in milliseconds 

        var dateCheck = sheet.getRange("C2:C50000").getValues()
        
        for (var x = 0 ; x < (dateCheck.length) ; x ++ ) {
          var dateCheckVal = new Date(dateCheck[x][0]).getTime() / 1000
          var dateCheckTweet = new Date(tweetDate).getTime() / 1000
          
    // Breaks the loop on blank values

          if (dateCheck[x][0] === "") {
            console.log("Breaking on blank-date check on entry #" + x)
            break
          } else if ( dateCheckVal == dateCheckTweet ) {

    // If it matches, tags it (true/false) and breaks the loop

            var tweetAlreadyExists = true
            console.log(dateCheck[x][0] + " matches on entry #" + x)
            break
          } else  { 
    
    // While no matches occur, keeps defining the tag as false

            var tweetAlreadyExists = false
            
          }
        }
        
    // Having this check done, the values are inserted in the spreadsheet

        if (!tweetAlreadyExists) { 
        
        sheet.getRange("A" + newRow).setValue(tweetID);
        sheet.getRange("A" + newRow).setNumberFormat("0000000000000000000");
        sheet.getRange("B" + newRow).setValue(tweetMsg);
        sheet.getRange("C" + newRow).setValue(tweetDate);
        sheet.getRange("C" + newRow).setNumberFormat("dd/MM/yyyy HH:MM:SS");
        sheet.getRange("D" + newRow).setValue(tweetRTCount);
        sheet.getRange("E" + newRow).setValue(tweetFavCount);
        nextRow = (nextRow + 1);
        console.log("Adding tweet #" + tweetID)
        }
    
    // If the Tweet already exists, no action is needed
    // The `if` statement will only add items that are not already in
    // the spreadsheet. So the `for` loop continues to cycle while 
    // there are still Tweets in the array.
    
      }
    }
  }
}  