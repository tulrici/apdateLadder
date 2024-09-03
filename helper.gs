// Function to find the row of a player in the ladder sheet
function findPlayerRow(playerName, sheet) {
  var range = sheet.getRange("C2:C" + sheet.getLastRow());
  var values = range.getValues();
  for (var i = 0; i < values.length; i++) {
    if (values[i][0].toString().toLowerCase() === playerName.toLowerCase()) {
      return i + 2;
    }
  }
  throw new Error("Player not found: " + playerName);
}

// Function to handle player name input and check for similar names
function handlePlayerName(playerName, sheet) {
  var similarPlayerName = checkSimilarPlayer(playerName, sheet);

  if (similarPlayerName === null) {
    createNewPlayer(playerName, sheet);
    return playerName;
  } else {
    if (similarPlayerName.toLowerCase() !== playerName.toLowerCase()) {
      var confirmUseSimilar = Browser.msgBox(
        "The name you entered '" + playerName + "' is too similar to '" + similarPlayerName + "'. Did you mean '" + similarPlayerName + "'?",
        Browser.Buttons.YES_NO
      );
      if (confirmUseSimilar === "yes") {
        return similarPlayerName; // Replace with the similar name
      }
    }
    return playerName; // Return the original name if no match or if user insists
  }
}

// Function to create a new player in the ladder sheet
function createNewPlayer(joueur, sheet) {
  var similarPlayerName = checkSimilarPlayer(joueur, sheet);
  if (similarPlayerName) {
    throw new Error("Player already exists or player name too similar to \"" + similarPlayerName + "\". Are you sure you didn't mean them?");
  }

  var lastRow = sheet.getLastRow() + 1;
  sheet.getRange("C" + lastRow).setValue(joueur);
  sheet.getRange("B" + lastRow).setValue(0); // Start at 0 points
  sheet.getRange("D" + lastRow + ":H" + lastRow).setValue(0); // Initialize all other values to 0
}

// Function to check for similar or duplicate player names in the correct sheet
function checkSimilarPlayer(newPlayerName, sheet) {
  var playerList = sheet.getRange("C2:C" + sheet.getLastRow()).getValues().flat();
  for (var i = 0; i < playerList.length; i++) {
    var existingPlayer = playerList[i];
    if (existingPlayer.toString().toLowerCase() === newPlayerName.toLowerCase()) {
      return existingPlayer; // Exact match
    }
    if (levenshteinDistance(existingPlayer.toString().toLowerCase(), newPlayerName.toLowerCase()) < 2) {
      return existingPlayer; // Similar name
    }
  }
  return null; // Return null if no match or similar name found
}

// Levenshtein Distance function to detect similar names (for typo prevention)
function levenshteinDistance(a, b) {
  var tmp;
  if (a.length === 0) { return b.length; }
  if (b.length === 0) { return a.length; }
  if (a.length > b.length) { tmp = a; a = b; b = tmp; }

  var i, j, res, alen = a.length, blen = b.length, row = Array(alen);
  for (i = 0; i <= alen; i++) { row[i] = i; }

  for (i = 1; i <= blen; i++) {
    res = i;
    for (j = 1; j <= alen; j++) {
      tmp = row[j - 1];
      row[j - 1] = res;
      res = b[i - 1] === a[j - 1] ? tmp : Math.min(tmp + 1, Math.min(res + 1, row[j] + 1));
    }
  }
  return res;
}

// Function to update game statistics in the ladder sheet
function updateGameStats(row, result, sheet) {
  var winCell = sheet.getRange("E" + row);
  var loseCell = sheet.getRange("F" + row);
  var drawCell = sheet.getRange("G" + row);
  if (result > 0) {
    winCell.setValue(winCell.getValue() + 1);
  } else if (result < 0) {
    loseCell.setValue(loseCell.getValue() + 1);
  } else {
    drawCell.setValue(drawCell.getValue() + 1);
  }
  var gamesPlayed = winCell.getValue() + loseCell.getValue() + drawCell.getValue();
  sheet.getRange("D" + row).setValue(gamesPlayed); // Update games played
}

// Function to calculate and update win percentage
function calculateWinPercentage(row, sheet) {
  var wins = sheet.getRange("E" + row).getValue();
  var gamesPlayed = sheet.getRange("D" + row).getValue();
  var winPercentage = gamesPlayed > 0 ? (wins / gamesPlayed) : 0;
  sheet.getRange("H" + row).setValue(winPercentage);
}

// Function to calculate and update average score
function calculateAverageScore(row, sheet, gameScore) {
  var averageScore = sheet.getRange("I" + row).getValue();
  var gamesPlayed = sheet.getRange("D" + row).getValue();
  var averageScore = gamesPlayed > 0 ? (averageScore + gameScore) / 2 : 0;
  sheet.getRange("I" + row).setValue(averageScore);
}

// Function to sort the ladder sheet based on points, wins, average score, games played, and alphabetical order
function sortAndRankLadder(ladderSheet) {
  // Sort the ladder sheet based on points, wins, average score, and games played
  var rangeToSort = ladderSheet.getRange(2, 1, ladderSheet.getLastRow() - 1, 9);
  rangeToSort.sort([
    {column: 2, ascending: false}, // Sort by "Points !" descending
    {column: 5, ascending: false}, // Then by "nb de win" descending
    {column: 9, ascending: false}, // Then by "average score" descending
    {column: 4, ascending: false}, // Then by "nb de parties" descending
    {column: 3, ascending: true}   // Finally by alphabetical order of player names
  ]);

  // Update the "Classement" column with the correct ranking
  for (var i = 0; i < rangeToSort.getNumRows(); i++) {
    ladderSheet.getRange(i + 2, 1).setValue(i + 1); // Write the rank in the "Classement" column
  }
}

// Function to add the game data to the "Games jouées" sheet
function addGameData(sheet, joueur1, joueur2, scoreJ1, scoreJ2, resultatJ1, resultatJ2, newPointsJ1, newPointsJ2) {
  // Find the last row that has data in column A
  var lastRow = sheet.getRange("A:A").getValues().filter(String).length + 1;

  // Determine the result for each player
  var resJ1 = resultatJ1 > 0 ? "win" : resultatJ1 < 0 ? "lose" : "Egal";
  var resJ2 = resultatJ2 > 0 ? "win" : resultatJ2 < 0 ? "lose" : "Egal";

  // Determine the last game number
  var lastGameNumber = lastRow > 2 ? sheet.getRange("A" + (lastRow - 1)).getValue() : 0;

  // If it's the first game of the season, start with game number 1
  var newGameNumber = lastGameNumber + 1;

  // Add the game data to the sheet
  sheet.getRange("A" + lastRow + ":J" + lastRow).setValues([[
    newGameNumber,
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy"),
    joueur1,
    joueur2,
    scoreJ1,
    scoreJ2,
    resJ1,
    resJ2,
    newPointsJ1,
    newPointsJ2
  ]]);
}