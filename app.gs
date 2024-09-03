function updateLadder() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ladderSheet = ss.getSheetByName("LADDER");
  var gamesSheet = ss.getSheetByName("Games jouées");

  if (!ladderSheet || !gamesSheet) {
    Browser.msgBox("Error: Could not find the required sheets.");
    return; // Exit if sheets are not found
  }

  // Check if the ladder is new (no players yet)
  var isNewLadder = ladderSheet.getLastRow() <= 1;

  // Track initial state to revert if necessary
  var initialLadderData = isNewLadder ? [] : ladderSheet.getRange(2, 1, ladderSheet.getLastRow() - 1, ladderSheet.getLastColumn()).getValues();

  try {
    // Ensure the sheets were found
    if (!ladderSheet || !gamesSheet) {
      throw new Error("Required sheets not found. Please check the sheet names.");
    }

    // Prompt user for Joueur 1 name
    var joueur1 = Browser.inputBox("Enter the name of Joueur 1:");
    if (!joueur1 || joueur1 === 'cancel') throw new Error("Process canceled by the user.");

    // Check for similar matches to Joueur 1's name using the correct sheet
    joueur1 = handlePlayerName(joueur1, ladderSheet);
    if (!joueur1) throw new Error("Process canceled by the user.");

    // Prompt user for Joueur 2 name
    var joueur2 = Browser.inputBox("Enter the name of Joueur 2:");
    if (!joueur2 || joueur2 === 'cancel') throw new Error("Process canceled by the user.");

    // Check for similar matches to Joueur 2's name using the correct sheet
    joueur2 = handlePlayerName(joueur2, ladderSheet);
    if (!joueur2) throw new Error("Process canceled by the user.");

    // Ensure Joueur 1 and Joueur 2 are not the same
    if (joueur1.toLowerCase() === joueur2.toLowerCase()) {
      throw new Error("Player 1 and Player 2 cannot be the same!");
    }

    // Prompt for Joueur 1's score (Joueur 2's score will be auto-calculated)
    var scoreJ1 = parseInt(Browser.inputBox("Enter Score for " + joueur1 + " (0-20):"), 10);
    if (isNaN(scoreJ1) || scoreJ1 < 0 || scoreJ1 > 20) {
      throw new Error("Invalid score for " + joueur1 + ". Must be an integer between 0 and 20.");
    }
    var scoreJ2 = 20 - scoreJ1;

    // Confirm Joueur 2's automatic score
    var confirmScore = Browser.msgBox(joueur2 + "'s score is automatically set to " + scoreJ2, Browser.Buttons.OK_CANCEL);
    if (confirmScore === 'cancel') throw new Error("Process canceled by the user.");

    // Find the rows for both players in the ladder
    var rowJ1 = findPlayerRow(joueur1, ladderSheet);
    var rowJ2 = findPlayerRow(joueur2, ladderSheet);

    // Calculate the base result based on the game score difference
    var resultatJ1 = scoreJ1 > scoreJ2 ? 10 : scoreJ1 < scoreJ2 ? -10 : 0;
    var resultatJ2 = -resultatJ1; // Opposite of Joueur 1's result

    // Calculate the ladder score difference between the two players
    var ladderScoreJ1 = ladderSheet.getRange("B" + rowJ1).getValue();
    var ladderScoreJ2 = ladderSheet.getRange("B" + rowJ2).getValue();
    var scoreDiff = Math.abs(ladderScoreJ1 - ladderScoreJ2); // Use ladder scores for the difference
    var additionalPoints = Math.floor(scoreDiff / 10);

    // Check if both players have the same ladder score
    if (ladderScoreJ1 !== ladderScoreJ2) {
      if (ladderScoreJ1 > ladderScoreJ2) {
        // Joueur 1 is higher-ranked
        resultatJ1 -= additionalPoints;
        resultatJ2 += additionalPoints;
      } else {
        // Joueur 2 is higher-ranked
        resultatJ1 += additionalPoints;
        resultatJ2 -= additionalPoints;
      }
    }

    // Update ladder scores for both players
    var newPointsJ1 = ladderScoreJ1 + resultatJ1;
    var newPointsJ2 = ladderScoreJ2 + resultatJ2;
    ladderSheet.getRange("B" + rowJ1).setValue(newPointsJ1);
    ladderSheet.getRange("B" + rowJ2).setValue(newPointsJ2);

// Generate the game recap with improved formatting
var recapMessage = "Game Recap\n" +
                   "-----------------------------------------\n" +
                   "Match: " + joueur1 + " vs " + joueur2 + "\n" +
                   "-----------------------------------------\n" +
                   "Scores:\n" +
                   "  - " + joueur1 + ": " + scoreJ1 + "\n" +
                   "  - " + joueur2 + ": " + scoreJ2 + "\n" +
                   "-----------------------------------------\n" +
                   "Result:\n" +
                   "  - " + joueur1 + ": " + (resultatJ1 > 0 ? "Win" : resultatJ1 < 0 ? "Lose" : "Draw") + "\n" +
                   "  - " + joueur2 + ": " + (resultatJ2 > 0 ? "Win" : resultatJ2 < 0 ? "Lose" : "Draw") + "\n" +
                   "-----------------------------------------\n" +
                   "New Points:\n" +
                   "  - " + joueur1 + ": " + newPointsJ1 + " pts\n" +
                   "  - " + joueur2 + ": " + newPointsJ2 + " pts\n" +
                   "-----------------------------------------\n" +
                   "Do you want to finalize this game entry?";

// Prompt user for confirmation
var confirmRecap = Browser.msgBox(recapMessage, Browser.Buttons.YES_NO);

if (confirmRecap === 'no') {
  throw new Error("Process canceled by the user.");
}
    // Update game statistics and other calculations for both players
    updateGameStats(rowJ1, resultatJ1, ladderSheet);
    updateGameStats(rowJ2, resultatJ2, ladderSheet);
    calculateWinPercentage(rowJ1, ladderSheet);
    calculateWinPercentage(rowJ2, ladderSheet);
    calculateAverageScore(rowJ1, ladderSheet, scoreJ1);
    calculateAverageScore(rowJ2, ladderSheet, scoreJ2);

    // Sort the ladder sheet based on points, wins, and average score
    sortAndRankLadder(ladderSheet);

    // Add the game data to the "Games jouées" sheet
    addGameData(gamesSheet, joueur1, joueur2, scoreJ1, scoreJ2, resultatJ1, resultatJ2, newPointsJ1, newPointsJ2);
        
    // Show a confirmation message
    Browser.msgBox("Ladder updated successfully!");

  } catch (e) {
    // Revert to the original state if an error occurs
    if (!isNewLadder) {
      ladderSheet.getRange(2, 1, initialLadderData.length, initialLadderData[0].length).setValues(initialLadderData);
    }
    Browser.msgBox("An error occurred or the process was canceled: " + e.message);
  }
}