/*
 * Wish Tally - Conversion
 * Version 1.5 made by yippy
 * https://github.com/Yippy/wish-tally-conversion-sheet
 */


function removeCache() {
  // Remove all cache
  var fileConvertedCache = DriveApp.getFilesByName(TEMP_SOURCE_TITLE);
  while (fileConvertedCache.hasNext()) {
    fileConvertedCache.next().setTrashed(true);
  }
}

function convertToWishTally(wishTallySource, cacheConvertedSource, autoImportSheet) {
  var genshinGachaExportSheet = SpreadsheetApp.getActive().getSheetByName(GENSHIN_GACHA_EXPORT_SHEET_NAME);
  
  if (cacheConvertedSource && wishTallySource && genshinGachaExportSheet) {
    for (const [key, value] of Object.entries(GENSHIN_GACHA_EXPORT_SHEET_NAMES_FROM_FILE)) {
      var isSkipped = true;
      var bannerSelection = autoImportSheet.getRange(RANGE_AUTO_IMPORT_SELECTION_BY_BANNER_NAMES[value]).getValue();
      
      if (bannerSelection) {
        isSkipped = false;
      }
      if (isSkipped) {
        autoImportSheet.getRange(RANGE_AUTO_IMPORT_STATUS_BY_BANNER_NAMES[value]).setValue("Skipped");
      } else {
        genshinGachaExportClearSheet();
        autoImportSheet.getRange(RANGE_AUTO_IMPORT_STATUS_BY_BANNER_NAMES[value]).setValue("Begin importing..");
        var wishCacheConvertedSheet = cacheConvertedSource.getSheetByName(key);
        var lastRowWithoutTitle = wishCacheConvertedSheet.getRange(2, 1, wishCacheConvertedSheet.getLastRow(), 1).getValues().filter(String).length;
        var wishTallySheet = wishTallySource.getSheetByName(value);
        var lastRowWithoutTitlewishTallySheet = wishTallySheet.getRange(2, 1, wishTallySheet.getLastRow(), 1).getValues().filter(String).length;
        var difference = lastRowWithoutTitle-lastRowWithoutTitlewishTallySheet;
        if (difference <= 0 || lastRowWithoutTitlewishTallySheet == lastRowWithoutTitle) {
          if (difference < 0){
            autoImportSheet.getRange(RANGE_AUTO_IMPORT_STATUS_BY_BANNER_NAMES[value]).setValue("Error - Wish Tally got more Wishes");
          } else {
            autoImportSheet.getRange(RANGE_AUTO_IMPORT_STATUS_BY_BANNER_NAMES[value]).setValue(difference+"/"+ lastRowWithoutTitle+" Nothing to import");
          }
        } else {
          autoImportSheet.getRange(RANGE_AUTO_IMPORT_STATUS_BY_BANNER_NAMES[value]).setValue("Found "+difference+" new wishes");
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
            autoImportSheet.getRange(RANGE_AUTO_IMPORT_STATUS_BY_BANNER_NAMES[value]).setValue(difference+"/"+ lastRowWithoutTitle+" Wishes added to banner");
          } else {
            autoImportSheet.getRange(RANGE_AUTO_IMPORT_STATUS_BY_BANNER_NAMES[value]).setValue(lastRowWithoutTitle+" Wishes imported");
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
  var autoImportSheet = SpreadsheetApp.getActive().getSheetByName(AUTO_IMPORT_SHEET_NAME);
  var genshinGachaExportFileType = autoImportSheet.getRange(RANGE_EXPORT_GENSHIN_GACHA_FILE_TYPE).getValue();
  var genshinGachaExportGoogleSheetFileType = autoImportSheet.getRange(RANGE_EXPORT_GENSHIN_GACHA_GOOGLE_SHEET_TYPE).getValue();
  if (autoImportSheet) {
    var fileTypeSelection = autoImportSheet.getRange(RANGE_FILE_TYPE_SELECTION).getValue();
    if (fileTypeSelection && fileTypeSelection == genshinGachaExportGoogleSheetFileType) {
      var sourceURL = autoImportSheet.getRange(RANGE_FILE_URL_USER_INPUT).getValue();
      var cacheConvertedSource = SpreadsheetApp.openByUrl(sourceURL);
      var wishTallyURL = autoImportSheet.getRange(RANGE_WISH_TALLY_URL_USER_INPUT).getValue();
      if (wishTallyURL != "") {
        var wishTallySource = SpreadsheetApp.openByUrl(wishTallyURL);
        convertToWishTally(wishTallySource, cacheConvertedSource, autoImportSheet);
      } else {
        title = "Error";
        message = "Must provide Wish Tally sheet URL, check cell "+RANGE_WISH_TALLY_URL_USER_INPUT;
      }
    } else if (fileTypeSelection && fileTypeSelection == genshinGachaExportFileType) {
      var fileID = getIdFromUrl(autoImportSheet.getRange(RANGE_FILE_URL_USER_INPUT).getValue());
      if (fileID) {
        var fileSource = DriveApp.getFileById(fileID);
        if (fileSource.getMimeType() == MimeType.MICROSOFT_EXCEL) {
          removeCache();

          var xBlob = fileSource.getBlob();
          var newFile = { title : TEMP_SOURCE_TITLE,
                         key : fileID
                        }
          var fileConvertedSource = Drive.Files.insert(newFile, xBlob, {
            convert: true
          });
          var wishTallyURL = autoImportSheet.getRange(RANGE_WISH_TALLY_URL_USER_INPUT).getValue();
          var cacheConvertedSource = SpreadsheetApp.openById(fileConvertedSource.getId());
          if (wishTallyURL != "") {
            var wishTallySource = SpreadsheetApp.openByUrl(wishTallyURL);
            convertToWishTally(wishTallySource, cacheConvertedSource,autoImportSheet);
          } else {
            title = "Error";
            message = "Must provide Wish Tally sheet URL, check cell "+RANGE_WISH_TALLY_URL_USER_INPUT;
          }
        } else {
          title = "Error";
          message = "Source is not an Excel file, check cell "+RANGE_WISH_TALLY_URL_USER_INPUT;
        }
      } else {
        title = "Error";
        message = "Must provide source file URL to import wishes, check cell "+RANGE_FILE_URL_USER_INPUT;
      }
    } else {
      title = "Error";
      message = "Must select a file type, check cell "+RANGE_FILE_TYPE_SELECTION;
    }
  } else {
    title = "Error";
    message = "Missing sheet named "+autoImportSheet;
  }
  if (title && title == "Error") {
     for (const [key, value] of Object.entries(RANGE_AUTO_IMPORT_STATUS_BY_BANNER_NAMES)) {
        autoImportSheet.getRange(RANGE_AUTO_IMPORT_STATUS_BY_BANNER_NAMES[key]).setValue("Failed");
     }
  }
  removeCache();
  SpreadsheetApp.getActiveSpreadsheet().toast(message, title);
}

function getIdFromUrl(url) {
  return url.match(/[-\w]{25,}/);
}

function exportClearSheet(exportFormat) {
  var exportSheet = SpreadsheetApp.getActive().getSheetByName(exportFormat);
  if (exportSheet) {
    var clearRows = exportSheet.getMaxRows()-2;
    if (clearRows > 0) {
      exportSheet.getRange(3, 1, clearRows, 6).clearContent();
    }
  } else {
    var title = "Error";
    var message = "Missing sheet named "+exportFormat;
    SpreadsheetApp.getActiveSpreadsheet().toast(message, title);
  }
}

function exportSortSheet(exportFormat) {
  var exportSheet = SpreadsheetApp.getActive().getSheetByName(exportFormat);
  if (exportSheet) {
    var lastRowWithoutTitle = exportSheet.getMaxRows()-2;
    var range = exportSheet.getRange(3, 1,lastRowWithoutTitle, 6);
    if (exportFormat == GENSHIN_WISHES_EXPORT_SHEET_NAME) {
      range.sort([{column: 3, ascending: true},{column: 4, ascending: true}]);
    } else {
      range.sort([{column: 5, ascending: true}]);
    }
  } else {
    var title = "Error";
    var message = "Missing sheet named "+exportFormat;
    SpreadsheetApp.getActiveSpreadsheet().toast(message, title);
  }
}

function exportAdjustFormat(exportFormat) {
  var exportSheet = SpreadsheetApp.getActive().getSheetByName(exportFormat);
  if (exportSheet) {
    var lastRowWithoutTitle = exportSheet.getMaxRows()-2;
    exportSheet.getRange(1, 1, exportSheet.getMaxRows(), exportSheet.getMaxColumns()).clearFormat();
    if (exportFormat == GENSHIN_WISHES_EXPORT_SHEET_NAME) {
      exportSheet.getRange(3, 3, lastRowWithoutTitle, 1).setBackground("lightyellow");
      exportSheet.getRange(3, 8, lastRowWithoutTitle, 4).setBackground("lightgrey");
    } else {
      exportSheet.getRange(3, 7, lastRowWithoutTitle, 4).setBackground("lightgrey");
    }
    exportSheet.getRange(1, 1, 2, 10).setNumberFormat("@");
    exportSheet.getRange(3, 1, lastRowWithoutTitle, 3).setNumberFormat("@");
    exportSheet.getRange(3, 4, lastRowWithoutTitle, 3).setNumberFormat("0");
    exportSheet.getRange(3, 7, lastRowWithoutTitle, 1).setNumberFormat("@");
    exportSheet.getRange(3, 8, lastRowWithoutTitle, 1).setNumberFormat("0");
    exportSheet.getRange(3, 9, lastRowWithoutTitle, 2).setNumberFormat("@");
  } else {
    var title = "Error";
    var message = "Missing sheet named "+exportFormat;
    SpreadsheetApp.getActiveSpreadsheet().toast(message, title);
  }
}

// Genshin Gacha
function genshinGachaExportAdjustAndSortSheet() {
  exportAdjustFormat(GENSHIN_GACHA_EXPORT_SHEET_NAME);
  exportSortSheet(GENSHIN_GACHA_EXPORT_SHEET_NAME);
}

function genshinGachaExportAdjustFormat() {
  exportAdjustFormat(GENSHIN_GACHA_EXPORT_SHEET_NAME);
}

function genshinGachaExportSortSheet() {
  exportSortSheet(GENSHIN_GACHA_EXPORT_SHEET_NAME);
}

function genshinGachaExportClearSheet() {
  exportClearSheet(GENSHIN_GACHA_EXPORT_SHEET_NAME);
}

// Genshin Wishes
function genshinWishesExportAdjustAndSortSheet() {
  exportAdjustFormat(GENSHIN_WISHES_EXPORT_SHEET_NAME);
  exportSortSheet(GENSHIN_WISHES_EXPORT_SHEET_NAME);
}

function genshinWishesExportAdjustFormat() {
  exportAdjustFormat(GENSHIN_WISHES_EXPORT_SHEET_NAME);
}

function genshinWishesExportSortSheet() {
  exportSortSheet(GENSHIN_WISHES_EXPORT_SHEET_NAME);
}

function genshinWishesExportClearSheet() {
  exportClearSheet(GENSHIN_WISHES_EXPORT_SHEET_NAME);
}