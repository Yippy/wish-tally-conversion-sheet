/*
 * Wish Tally - Conversion
 * Version 1.4 made by yippym
 */


function onOpen( ){
  var ui = SpreadsheetApp.getUi();
  
  ui.createMenu('Wish Tally')
  .addSeparator()
  .addSubMenu(ui.createMenu('genshin-gacha-export')
             .addItem('Clear', 'genshinGachaExportClearSheet')
             .addItem('Adjust Format', 'genshinGachaExportAdjustFormat')
             .addItem('Sort Wish Count', 'genshinGachaExportSortSheet')
             .addItem('Adjust and Sort', 'genshinGachaExportAdjustAndSortSheet'))
  .addSeparator()
  .addItem('Auto Import', 'autoImportToWishTally')
  .addSeparator()
  .addToUi();
}

var autoImportSheetName = "Auto Import";
var genshinGachaExportSheetName = "genshin-gacha-export";
var genshinGachaExportFileType = "genshin-gacha-export xlsx file";
var genshinGachaExportGoogleSheetFileType = "genshin-gacha-export xlsx converted to Google Sheet file";
var rangeFileType = "A4";
var rangeFileURL = "A7";
var rangeWishTallyURL = "A10";
var rangeCharacterEventWishHistoryBannerSelection = "B13";
var rangeWeaponEventWishHistoryBannerSelection = "B14";
var rangePermanentWishHistoryBannerSelection = "B15";
var rangeNoviceWishHistoryBannerSelection = "B16";
var rangeCharacterEventWishHistory = "B23";
var rangeWeaponEventWishHistory = "B24";
var rangePermanentWishHistory = "B25";
var rangeNoviceWishHistory = "B26";
var rangeDateStart = "B27";
var rangeDateEnd = "B28";
var genshinGachaExportSheetNamesFromFile = {
  "Character Event Wish": "Character Event Wish History",
  "Weapon Event Wish": "Weapon Event Wish History",
  "Permanent Wish": "Permanent Wish History",
  "Novice Wishes": "Novice Wish History"
}
var rangeAutoImportStatusByBannerNames = {
  "Character Event Wish History": rangeCharacterEventWishHistory,
  "Weapon Event Wish History": rangeWeaponEventWishHistory,
  "Permanent Wish History": rangePermanentWishHistory,
  "Novice Wish History": rangeNoviceWishHistory
}
var rangeAutoImportSelectionByBannerNames = {
  "Character Event Wish History": rangeCharacterEventWishHistoryBannerSelection,
  "Weapon Event Wish History": rangeWeaponEventWishHistoryBannerSelection,
  "Permanent Wish History": rangePermanentWishHistoryBannerSelection,
  "Novice Wish History": rangeNoviceWishHistoryBannerSelection
}

var tempSourceTitle = "GachaExport_WishTally_Conversion_Temp";
function removeCache() {
  // Remove all cache
  var fileConvertedCache = DriveApp.getFilesByName(tempSourceTitle);
  while (fileConvertedCache.hasNext()) {
    fileConvertedCache.next().setTrashed(true);
  }
}


function convertToWishTally(wishTallySource, cacheConvertedSource, autoImportSheet) {
  var genshinGachaExportSheet = SpreadsheetApp.getActive().getSheetByName(genshinGachaExportSheetName);
  
  if (cacheConvertedSource && wishTallySource && genshinGachaExportSheet) {
    for (const [key, value] of Object.entries(genshinGachaExportSheetNamesFromFile)) {
      var isSkipped = true;
      var bannerSelection = autoImportSheet.getRange(rangeAutoImportSelectionByBannerNames[value]).getValue();
      
      if (bannerSelection) {
        isSkipped = false;
      }
      if (isSkipped) {
        autoImportSheet.getRange(rangeAutoImportStatusByBannerNames[value]).setValue("Skipped");
      } else {
        genshinGachaExportClearSheet();
        autoImportSheet.getRange(rangeAutoImportStatusByBannerNames[value]).setValue("Begin importing..");
        var wishCacheConvertedSheet = cacheConvertedSource.getSheetByName(key);
        var lastRowWithoutTitle = wishCacheConvertedSheet.getRange(2, 1, wishCacheConvertedSheet.getLastRow(), 1).getValues().filter(String).length;
        var wishTallySheet = wishTallySource.getSheetByName(value);
        var lastRowWithoutTitlewishTallySheet = wishTallySheet.getRange(2, 1, wishTallySheet.getLastRow(), 1).getValues().filter(String).length;
        var difference = lastRowWithoutTitle-lastRowWithoutTitlewishTallySheet;
        if (difference <= 0 || lastRowWithoutTitlewishTallySheet == lastRowWithoutTitle) {
          if (difference < 0){
            autoImportSheet.getRange(rangeAutoImportStatusByBannerNames[value]).setValue("Error - Wish Tally got more Wishes");
          } else {
            autoImportSheet.getRange(rangeAutoImportStatusByBannerNames[value]).setValue(difference+"/"+ lastRowWithoutTitle+" Nothing to import");
          }
        } else {
          autoImportSheet.getRange(rangeAutoImportStatusByBannerNames[value]).setValue("Found "+difference+" new wishes");
          var fileSourceRange = wishCacheConvertedSheet.getRange(2+(lastRowWithoutTitle-difference), 1, difference, 6).getValues();
          genshinGachaExportSheet.getRange(3, 1, difference, 6).setValues(fileSourceRange);
          
          SpreadsheetApp.getActiveSpreadsheet().toast("Converting "+difference+" wishes", key);
          SpreadsheetApp.flush();
          //Give time for sheet to sort array formula
          Utilities.sleep(10*1000);
          SpreadsheetApp.flush();
          var wishTallyConvertRange = genshinGachaExportSheet.getRange(3, 7, difference, 2).getValues();
          wishTallySheet.getRange(2+lastRowWithoutTitlewishTallySheet, 1, difference, 2).setValues(wishTallyConvertRange);
          if (lastRowWithoutTitlewishTallySheet > 0) {
            autoImportSheet.getRange(rangeAutoImportStatusByBannerNames[value]).setValue(difference+"/"+ lastRowWithoutTitle+" Wishes added to banner");
          } else {
            autoImportSheet.getRange(rangeAutoImportStatusByBannerNames[value]).setValue(lastRowWithoutTitle+" Wishes imported");
          }
        }
      }
    }
    title = "Complete";
    message = "";
  } else {
    title = "Error";
    message = "Unable to load wish tally";
  }
}

function autoImportToWishTally() {
  var message = "";
  var title = "";
  var autoImportSheet = SpreadsheetApp.getActive().getSheetByName(autoImportSheetName);
  
  if (autoImportSheet) {
    var fileTypeSelection = autoImportSheet.getRange(rangeFileType).getValue();
    if (fileTypeSelection && fileTypeSelection == genshinGachaExportGoogleSheetFileType) {
      var sourceURL = autoImportSheet.getRange(rangeFileURL).getValue();
      var cacheConvertedSource = SpreadsheetApp.openByUrl(sourceURL);
      var wishTallyURL = autoImportSheet.getRange(rangeWishTallyURL).getValue();
      if (wishTallyURL != "") {
        var wishTallySource = SpreadsheetApp.openByUrl(wishTallyURL);
        convertToWishTally(wishTallySource, cacheConvertedSource, autoImportSheet);
      } else {
        title = "Error";
        message = "Must provide Wish Tally sheet URL, check cell "+rangeWishTallyURL;
      }
    } else if (fileTypeSelection && fileTypeSelection == genshinGachaExportFileType) {
      var fileID = getIdFromUrl(autoImportSheet.getRange(rangeFileURL).getValue());
      if (fileID) {
        var fileSource = DriveApp.getFileById(fileID);
        if (fileSource.getMimeType() == MimeType.MICROSOFT_EXCEL) {
          removeCache();

          var xBlob = fileSource.getBlob();
          var newFile = { title : tempSourceTitle,
                         key : fileID
                        }
          var fileConvertedSource = Drive.Files.insert(newFile, xBlob, {
            convert: true
          });
          var wishTallyURL = autoImportSheet.getRange(rangeWishTallyURL).getValue();
          var cacheConvertedSource = SpreadsheetApp.openById(fileConvertedSource.getId());
          if (wishTallyURL != "") {
            var wishTallySource = SpreadsheetApp.openByUrl(wishTallyURL);
            convertToWishTally(wishTallySource, cacheConvertedSource,autoImportSheet);
          } else {
            title = "Error";
            message = "Must provide Wish Tally sheet URL, check cell "+rangeWishTallyURL;
          }
        } else {
          title = "Error";
          message = "Source is not an Excel file, check cell "+rangeWishTallyURL;
        }
      } else {
        title = "Error";
        message = "Must provide source file URL to import wishes, check cell "+rangeFileURL;
      }
    } else {
      title = "Error";
      message = "Must select a file type, check cell "+rangeFileType;
    }
  } else {
    title = "Error";
    message = "Missing sheet named "+autoImportSheet;
  }
  if (title && title == "Error") {
     for (const [key, value] of Object.entries(rangeAutoImportStatusByBannerNames)) {
        autoImportSheet.getRange(rangeAutoImportStatusByBannerNames[key]).setValue("Failed");
     }
  }
  removeCache();
  SpreadsheetApp.getActiveSpreadsheet().toast(message, title);
}

function getIdFromUrl(url) {
  return url.match(/[-\w]{25,}/);
}

function genshinGachaExportClearSheet() {
  var genshinGachaExportSheet = SpreadsheetApp.getActive().getSheetByName(genshinGachaExportSheetName);
  if (genshinGachaExportSheet) {
    var clearRows = genshinGachaExportSheet.getMaxRows()-2;
    if (clearRows > 0) {
      genshinGachaExportSheet.getRange(3, 1, clearRows, 6).clearContent();
    }
  } else {
    var title = "Error";
    var message = "Missing sheet named "+genshinGachaExportSheetName;
    SpreadsheetApp.getActiveSpreadsheet().toast(message, title);
  }
}

function genshinGachaExportSortSheet() {
  var genshinGachaExportSheet = SpreadsheetApp.getActive().getSheetByName(genshinGachaExportSheetName);
  if (genshinGachaExportSheet) {
    var lastRowWithoutTitle = genshinGachaExportSheet.getMaxRows()-2;
    var range = genshinGachaExportSheet.getRange(3, 1,lastRowWithoutTitle, 6);
    range.sort([{column: 5, ascending: true}]);
  } else {
    var title = "Error";
    var message = "Missing sheet named "+genshinGachaExportSheetName;
    SpreadsheetApp.getActiveSpreadsheet().toast(message, title);
  }
}

function genshinGachaExportAdjustFormat() {
  var genshinGachaExportSheet = SpreadsheetApp.getActive().getSheetByName(genshinGachaExportSheetName);
  if (genshinGachaExportSheet) {
    var lastRowWithoutTitle = genshinGachaExportSheet.getMaxRows()-2;
    genshinGachaExportSheet.getRange(1, 1, genshinGachaExportSheet.getMaxRows(), genshinGachaExportSheet.getMaxColumns()).clearFormat();
    genshinGachaExportSheet.getRange(3, 7,lastRowWithoutTitle, 4).setBackground("lightgrey");
    genshinGachaExportSheet.getRange(1, 1, 2, 10).setNumberFormat("@");
    genshinGachaExportSheet.getRange(3, 1, lastRowWithoutTitle, 3).setNumberFormat("@");
    genshinGachaExportSheet.getRange(3, 4, lastRowWithoutTitle, 3).setNumberFormat("0");
    genshinGachaExportSheet.getRange(3, 7, lastRowWithoutTitle, 1).setNumberFormat("@");
    genshinGachaExportSheet.getRange(3, 8, lastRowWithoutTitle, 1).setNumberFormat("0");
    genshinGachaExportSheet.getRange(3, 9, lastRowWithoutTitle, 2).setNumberFormat("@");
  } else {
    var title = "Error";
    var message = "Missing sheet named "+genshinGachaExportSheetName;
    SpreadsheetApp.getActiveSpreadsheet().toast(message, title);
  }
}

function genshinGachaExportAdjustAndSortSheet() {
  genshinGachaExportAdjustFormat();
  genshinGachaExportSortSheet();
  
}