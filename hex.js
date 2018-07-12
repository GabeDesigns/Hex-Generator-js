//Global Variables

var sheet = SpreadsheetApp.getActiveSpreadsheet();
var numOfCombos;
var schoolChosen;
var comboArray = new Array();
var hexValues = new Array();

function getValues() {
  var sheetName;
  var x = 0;

  for (var i = 0; i < 4; ++i) {
    //Browser.msgBox(cobbCValues[i]);
    if (i == 0) {
      sheetName = "Cobb";
    } else if (i == 1) {
      sheetName = "Hamblen";
    } else if (i == 2) {
      sheetName = "McMullan";
    } else if (i == 3) {
      sheetName = "Brown";
    } else {
    }

    var activeSheet = sheet.getSheetByName(sheetName);
    var activeCRange = activeSheet.getRange(
      2,
      1,
      activeSheet.getLastRow() - 1,
      1
    );
    var activeCValues = activeCRange.getValues();

    for (var h = 0; activeCValues[h] != null; h += 5) {
      comboArray[x] = activeCValues[h];
      ++x;
    }
  }

  //Browser.msgBox(comboArray);

  var generateSheet = sheet.getSheetByName("Generate");
  var generateRange = generateSheet.getRange(4, 3);
  numOfCombos = generateRange.getValues();
  var schoolRange = generateSheet.getRange(6, 3);
  schoolChosen = schoolRange.getValues();

  var hexSheet = sheet.getSheetByName(schoolChosen);
  var hexRange = hexSheet.getRange(2, 2, hexSheet.getLastRow() - 1, 1);
  var hexValuesRaw = hexRange.getValues();

  for (i = 0; i < hexValuesRaw.length; ++i) {
    hexValues[i] = hexValuesRaw[i];
  }
  //Browser.msgBox(hexValues);

  generateCombos();
}

function generateCombos() {
  var first;
  var hexcode = "";
  var hex = "ABCDEF0123456789";
  var Base10 = "0123456789";
  var combo = "";
  var dup = false;

  var comboSheet = sheet.getSheetByName(schoolChosen);
  var comboRange = comboSheet.getRange(2, 1, comboSheet.getLastRow() - 1, 1);
  var comboValues = comboRange.getValues();

  var comboLength = comboValues.length;
  var rem = comboLength % 5;

  Browser.msgBox(rem);
  rem = 5 - rem;
  var idc = numOfCombos - rem;
  if (idc < 0) {
    rem = numOfCombos;
  }
  // run this:

  if (rem != 0) {
    //pull out of if

    var repeat = comboValues[comboLength - 1];
    for (var z = 0; z < rem; ++z) {
      comboSheet.getRange(comboSheet.getLastRow() + 1, 1).setValue(repeat);
    }
  }

  //  for(var i = 0; i < 4; i++)
  //  {
  //
  //   combo += Base10.charAt(Math.floor(Math.random() * Base10.length));
  //
  //  }
  //
  //  Browser.msgBox(combo);

  if (schoolChosen == "Cobb") {
    first = 0;
  } else if (schoolChosen == "Hamblen") {
    first = 1;
  } else if (schoolChosen == "McMullan") {
    first = 2;
  } else if (schoolChosen == "Brown") {
    first = 3;
  } else {
  }

  for (var q = 0; q <= numOfCombos; ++q) {
    hexcode = "";
    for (var i = 0; i < 3; i++) {
      hexcode += hex.charAt(Math.floor(Math.random() * hex.length));
    }
    hexcode = first + hexcode;
    //Browser.msgBox(hexcode);

    for (var g = 0; g <= hexValues.length; ++g) {
      if (hexcode == hexValues[g]) {
        //Browser.msgBox("dup");
        dup = true;
        break;
      } else {
      }
    }
    if (dup == false) {
      var w = hexValues.length;
      hexValues[w] = hexcode;
      //Browser.msgBox(hexValues);
      numOfCombos -= 1;
      // Browser.msgBox(numOfCombos);
    } else {
      break;
    }
  }

  if (dup == true) {
    generateCombos();
  } else {
  }
}
