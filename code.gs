function rankValue(rank) {
  var tierOrder = {
    "IRON": 0,
    "BRONZE": 1,
    "SILVER": 2,
    "GOLD": 3,
    "PLATINUM": 4,
    "EMERALD": 5,
    "DIAMOND": 6,
  };

  var divisionOrder = {
    "I" : 3,
    "II" : 2,
    "III" : 1,
    "IV" : 0
  };

  // Split rank into tier and division
  var parts = rank.split(" ");
  var tier = parts[0].toUpperCase();
  var division = parts[1];

  // Calculate the rank value
  var value = (tierOrder[tier] || 0) * 400 + (divisionOrder[division] || 0) * 100;
  if (parts.length > 2) { // If LP is present
    value += parseInt(parts[2]);
  }
  return value;
}

function sortData(sheet, startRow, endRow) {
  var dataRange = sheet.getRange(startRow, 1, endRow - startRow + 1, 8); // Assuming 6 columns of data
  var values = dataRange.getValues();
  values.sort(function(a, b) {
    return rankValue(b[2]) - rankValue(a[2]); // Reverse the order of comparison
  });
  dataRange.setValues(values);
}

function findSoloQueueIndex(thirdData) {
  for (var i = 0; i < thirdData.length; i++) {
    if (thirdData[i].queueType === "RANKED_SOLO_5x5") {
      return i; // Return the index if found
    }
  }
  return -1; // Return -1 if not found
}

function fetchRiotData() {
  // Get values from specific cells
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var namesRange = sheet.getRange("A3:A"); // Assuming names are in column A starting from row 3
  var taglinesRange = sheet.getRange("B3:B"); // Assuming taglines are in column B starting from row 3
  var api_key = "" // API_key

  // Get the values from the ranges
  var names = namesRange.getValues().flat().filter(String); // Flatten and remove empty strings
  var taglines = taglinesRange.getValues().flat().filter(String); // Flatten and remove empty strings

  // Ensure the lengths of names and taglines are the same
  if (names.length !== taglines.length) {
    Logger.log("Number of names and taglines don't match.");
    return;
  }

  // Loop through each name and tagline
  for (var i = 0; i < names.length; i++) {
    var name = names[i];
    var tagline = taglines[i];

    // Construct the first API URL to get puuid
    var firstApiUrl = "https://americas.api.riotgames.com/riot/account/v1/accounts/by-riot-id/" + name + "/" + tagline + "?api_key=" + api_key;

    try {
      // Fetch the JSON data
      var response = UrlFetchApp.fetch(firstApiUrl);
      var json = response.getContentText();
      var data = JSON.parse(json);

      // Extract puuid from the response
      var puuid = data.puuid;

      // Construct the second API URL using the retrieved puuid
      var secondApiUrl = "https://na1.api.riotgames.com/lol/summoner/v4/summoners/by-puuid/" + puuid + "?api_key=" + api_key;

      // Fetch more data using the puuid
      var secondResponse = UrlFetchApp.fetch(secondApiUrl);
      var secondJson = secondResponse.getContentText();
      var secondData = JSON.parse(secondJson);

      // Extract id from the secondData
      var id = secondData.id;

      // Construct the third API URL using the retrieved id
      var thirdApiUrl = "https://na1.api.riotgames.com/lol/league/v4/entries/by-summoner/" + id + "?api_key=" + api_key;

      // Fetch more data using the id
      var thirdResponse = UrlFetchApp.fetch(thirdApiUrl);
      var thirdJson = thirdResponse.getContentText();
      var thirdData = JSON.parse(thirdJson);

      var n = findSoloQueueIndex(thirdData)

      // Output the data to specific cells
      var row = 3 + i; // Start from row 3
      sheet.getRange("C" + row).setValue(thirdData[n].tier + " " + thirdData[n].rank + " " + thirdData[n].leaguePoints); // assuming tier is in the first element of the array
      sheet.getRange("D" + row).setValue(thirdData[n].wins); // assuming wins is in the first element of the array
      sheet.getRange("E" + row).setValue(thirdData[n].losses); // assuming losses is in the first element of the array
      var winrate = parseFloat(thirdData[n].wins) / (parseFloat(thirdData[n].wins) + parseFloat(thirdData[n].losses));
      sheet.getRange("F" + row).setValue(winrate).setNumberFormat("0.00%"); // set as percentage format with 2 decimal places

      if (rankValue(sheet.getRange("C" + row).getValue()) > rankValue(sheet.getRange("G" + row).getValue())) {
        sheet.getRange("G" + row).setValue(sheet.getRange("C" + row).getValue());
      }

      Utilities.sleep(4000); // 4000 milliseconds = 4 seconds To keep rates limited
    } catch (error) {
      // Handle errors by filling data with N/A
      var row = 3 + i; // Start from row 3
      sheet.getRange("C" + row).setValue("N/A");
      sheet.getRange("D" + row).setValue("N/A");
      sheet.getRange("E" + row).setValue("N/A");
      sheet.getRange("F" + row).setValue("N/A");
      Logger.log("Error processing data for name: " + name + " and tagline: " + tagline + ". Error message: " + error.message);

      Utilities.sleep(4000); // 4000 milliseconds = 4 seconds
    }
  }

  sortData(sheet, 3, 3 + names.length - 1);
}
