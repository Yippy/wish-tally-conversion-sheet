/*
 * Wish Tally - Conversion
 * Version 1.7 made by yippy
 * https://github.com/Yippy/wish-tally-conversion-sheet
 */


function removeCache() {
  // Remove all cache
  var fileConvertedCache = DriveApp.getFilesByName(TEMP_SOURCE_TITLE);
  while (fileConvertedCache.hasNext()) {
    fileConvertedCache.next().setTrashed(true);
  }
}

function convertGenshinGachaExportToWishTally(wishTallySource, cacheConvertedSource, autoImportSheet) {
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

function convertPaimonMoeExportToWishTally(wishTallySource, cacheConvertedSource, autoImportSheet) {
  var genshinGachaExportSheet = SpreadsheetApp.getActive().getSheetByName(PAIMON_MOE_EXPORT_SHEET_NAME);
  
  if (cacheConvertedSource && wishTallySource && genshinGachaExportSheet) {
    for (const [key, value] of Object.entries(PAIMON_MOE_EXPORT_SHEET_NAMES_FROM_FILE)) {
      var isSkipped = true;
      var bannerSelection = autoImportSheet.getRange(RANGE_AUTO_IMPORT_SELECTION_BY_BANNER_NAMES[value]).getValue();

      if (bannerSelection) {
        isSkipped = false;
      }
      if (isSkipped) {
        autoImportSheet.getRange(RANGE_AUTO_IMPORT_STATUS_BY_BANNER_NAMES[value]).setValue("Skipped");
      } else {
        paimonMoeExportClearSheet();
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
          var fileSourceRange = wishCacheConvertedSheet.getRange(2+(lastRowWithoutTitle-difference), 1, difference, 8).getValues();
          genshinGachaExportSheet.getRange(3, 1, difference, 8).setValues(fileSourceRange);
          
          SpreadsheetApp.getActiveSpreadsheet().toast("Converting "+difference+" wishes", key);
          SpreadsheetApp.flush();
          //Give time for sheet to sort array formula
          Utilities.sleep(10*1000);
          SpreadsheetApp.flush();
          var wishTallyConvertRange = genshinGachaExportSheet.getRange(3, 9, difference, 2).getValues();
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

function convertGenshinWishesExportToWishTally(wishTallySource, cacheConvertedSource, autoImportSheet) {
  var genshinWishesExportSheet = SpreadsheetApp.getActive().getSheetByName(GENSHIN_WISHES_EXPORT_SHEET_NAME);
  
  if (cacheConvertedSource && wishTallySource && genshinWishesExportSheet) {
    var wishes = [];
    var isSkipped = [];
    for (const [key, value] of Object.entries(GENSHIN_WISHES_EXPORT_SHEET_NAMES_FROM_FILE)) {
      wishes[key] = [];
      var isSkipped = true;
      var bannerSelection = autoImportSheet.getRange(RANGE_AUTO_IMPORT_SELECTION_BY_BANNER_NAMES[value]).getValue();
      if (bannerSelection) {
        isSkipped = false;
      }
      isSkipped[key] = isSkipped;
      if (isSkipped) {
        autoImportSheet.getRange(RANGE_AUTO_IMPORT_STATUS_BY_BANNER_NAMES[value]).setValue("Skipped");
      }
    }

    genshinWishesExportClearSheet();
    var wishCacheConvertedSheets = cacheConvertedSource.getSheets();
    var wishCacheConvertedSheet;
    // There should only be one sheet
    if (wishCacheConvertedSheets.length > 0) {
      for (var i = 0; i < wishCacheConvertedSheets.length; i++) {
        var sheetCheck = wishCacheConvertedSheets[i];
        if (GENSHIN_WISHES_HEADER == sheetCheck.getRange(GENSHIN_WISHES_HEADER_RANGE).getValue()) {
          wishCacheConvertedSheet = sheetCheck;
        }
      }
    }
    if (wishCacheConvertedSheet) {
      var lastRowWithoutTitle = wishCacheConvertedSheet.getRange(2, 1, wishCacheConvertedSheet.getLastRow(), 1).getValues().filter(String).length;

      var fileSourceRange = wishCacheConvertedSheet.getRange(2, 1, lastRowWithoutTitle, 1).getValues();
      genshinWishesExportSheet.getRange(3, 1, fileSourceRange.length, 1).setValues(fileSourceRange);
      SpreadsheetApp.flush();
      //Give time for sheet to sort array formula
      Utilities.sleep(10*1000);
      SpreadsheetApp.flush();

      var extractPasteValues = genshinWishesExportSheet.getRange(3, 8, fileSourceRange.length, 2).getValues();
      var extractBannerValues = genshinWishesExportSheet.getRange(3, 3, fileSourceRange.length, 1).getValues();
      // Sort wish history into banners
      for (var i = 0; i < extractBannerValues.length; i++) {
        var bannerName = extractBannerValues[i];
        var wishHistory = wishes[bannerName];
        wishHistory.push(extractPasteValues[i]);
        wishes[bannerName] = wishHistory;
      }
      for (const [key, value] of Object.entries(GENSHIN_WISHES_EXPORT_SHEET_NAMES_FROM_FILE)) {
        if (!isSkipped[key]) {
          var wishHistory = wishes[key];
          var lastRowWithoutTitle = wishHistory.length;
          var wishTallySheet = wishTallySource.getSheetByName(value);
          var lastRowWithoutTitlewishTallySheet = wishTallySheet.getRange(2, 1, wishTallySheet.getLastRow(), 1).getValues().filter(String).length;
          var difference = lastRowWithoutTitle-lastRowWithoutTitlewishTallySheet;

          if (difference <= 0 || lastRowWithoutTitlewishTallySheet ==  lastRowWithoutTitle) {
            if (difference < 0){
              autoImportSheet.getRange(RANGE_AUTO_IMPORT_STATUS_BY_BANNER_NAMES[value]).setValue("Error - Wish Tally got more Wishes");
            } else {
              autoImportSheet.getRange(RANGE_AUTO_IMPORT_STATUS_BY_BANNER_NAMES[value]).setValue(difference+"/"+ lastRowWithoutTitle+" Nothing to import");
            }
          } else {
            autoImportSheet.getRange(RANGE_AUTO_IMPORT_STATUS_BY_BANNER_NAMES[value]).setValue("Found "+difference+" new wishes");
            wishHistory = wishHistory.slice(lastRowWithoutTitlewishTallySheet);
            
            SpreadsheetApp.getActiveSpreadsheet().toast("Converting "+difference+" wishes", key);
            SpreadsheetApp.flush();
            //Give time to show status
            Utilities.sleep(10*1000);
            SpreadsheetApp.flush();
            wishTallySheet.getRange(2+lastRowWithoutTitlewishTallySheet, 1, difference, 2).setValues(wishHistory);
            if (lastRowWithoutTitlewishTallySheet > 0) {
              autoImportSheet.getRange(RANGE_AUTO_IMPORT_STATUS_BY_BANNER_NAMES[value]).setValue(difference+"/"+ lastRowWithoutTitle+" Wishes added to banner");
            } else {
              autoImportSheet.getRange(RANGE_AUTO_IMPORT_STATUS_BY_BANNER_NAMES[value]).setValue(lastRowWithoutTitle+" Wishes imported");
            }
          }
        }
      }
    } else {
      title = "Error";
      message = "Source does not have Genshin Wishes template";
      SpreadsheetApp.getActiveSpreadsheet().toast(message, title);
    }
    genshinWishesExportClearSheet();
  } else {
    title = "Error";
    message = "Unable to load wish tally";
    SpreadsheetApp.getActiveSpreadsheet().toast(message, title);
  }
}

function autoImportToWishTally() {
  var message = "";
  var title = "";
  var autoImportSheet = SpreadsheetApp.getActive().getSheetByName(AUTO_IMPORT_SHEET_NAME);
  var genshinGachaExportFileType = autoImportSheet.getRange(RANGE_EXPORT_GENSHIN_GACHA_FILE_TYPE).getValue();
  var genshinGachaExportGoogleSheetFileType = autoImportSheet.getRange(RANGE_EXPORT_GENSHIN_GACHA_GOOGLE_SHEET_TYPE).getValue();
  var genshinWishesExportFileType = autoImportSheet.getRange(RANGE_EXPORT_GENSHIN_WISHES_FILE_TYPE).getValue();
  var paimonMoeExportFileType = autoImportSheet.getRange(RANGE_EXPORT_PAIMON_MOE_FILE_TYPE).getValue();
  var paimonMoeExportGoogleSheetFileType = autoImportSheet.getRange(RANGE_EXPORT_PAIMON_MOE_GOOGLE_SHEET_TYPE).getValue();
  if (autoImportSheet) {
    var fileTypeSelection = autoImportSheet.getRange(RANGE_FILE_TYPE_SELECTION).getValue();
    if (fileTypeSelection) {
      if (fileTypeSelection == genshinGachaExportGoogleSheetFileType || fileTypeSelection == paimonMoeExportGoogleSheetFileType) {
        var sourceURL = autoImportSheet.getRange(RANGE_FILE_URL_USER_INPUT).getValue();
        var cacheConvertedSource;
        try {
          cacheConvertedSource = SpreadsheetApp.openByUrl(sourceURL);
        } catch(e) {
          title = "Error";
          message = "Invalid URL, check cell "+RANGE_FILE_URL_USER_INPUT+". Make sure the link is an Google Sheet";
          SpreadsheetApp.getActiveSpreadsheet().toast(message, title);
          return;
        }
        var wishTallyURL = autoImportSheet.getRange(RANGE_WISH_TALLY_URL_USER_INPUT).getValue();
        if (wishTallyURL != "") {
          var wishTallySource = SpreadsheetApp.openByUrl(wishTallyURL);
          if (fileTypeSelection == paimonMoeExportGoogleSheetFileType) {
            convertPaimonMoeExportToWishTally(wishTallySource, cacheConvertedSource,autoImportSheet);
          } else if (fileTypeSelection == genshinGachaExportGoogleSheetFileType) {
            convertGenshinGachaExportToWishTally(wishTallySource, cacheConvertedSource, autoImportSheet);
          } else {
            convertGenshinWishesExportToWishTally(wishTallySource, cacheConvertedSource,autoImportSheet);
          }
        } else {
          title = "Error";
          message = "Must provide Wish Tally sheet URL, check cell "+RANGE_WISH_TALLY_URL_USER_INPUT;
        }
      } else if (fileTypeSelection == genshinGachaExportFileType || fileTypeSelection == genshinWishesExportFileType || fileTypeSelection == paimonMoeExportFileType) {
        var sourceURL = autoImportSheet.getRange(RANGE_FILE_URL_USER_INPUT).getValue();
        var fileID = getIdFromUrl(sourceURL);
        if (fileID) {
          var fileSource = DriveApp.getFileById(fileID);
          var isValid = false;
          if (fileSource.getMimeType() == MimeType.MICROSOFT_EXCEL && fileTypeSelection == genshinGachaExportFileType) {
            isValid = true;
          } else if (fileSource.getMimeType() == MimeType.MICROSOFT_EXCEL && fileTypeSelection == paimonMoeExportFileType) {
            isValid = true;
          } else if (fileSource.getMimeType() == MimeType.CSV && fileTypeSelection == genshinWishesExportFileType) {
            isValid = true;
          }
          if (isValid) {
            removeCache();

            var xBlob = fileSource.getBlob();
            var newFile = { title : TEMP_SOURCE_TITLE,
                          key : fileID,
                          mimeType: MimeType.GOOGLE_SHEETS
                          }
            var fileConvertedSource = Drive.Files.insert(newFile, xBlob, {
              convert: true
            });
            var wishTallyURL = autoImportSheet.getRange(RANGE_WISH_TALLY_URL_USER_INPUT).getValue();
            var cacheConvertedSource = SpreadsheetApp.openById(fileConvertedSource.getId());
            if (wishTallyURL != "") {
              var wishTallySource = SpreadsheetApp.openByUrl(wishTallyURL);
              if (fileTypeSelection == paimonMoeExportFileType) {
                convertPaimonMoeExportToWishTally(wishTallySource, cacheConvertedSource,autoImportSheet);
              } else if (fileTypeSelection == genshinGachaExportFileType) {
                convertGenshinGachaExportToWishTally(wishTallySource, cacheConvertedSource,autoImportSheet);
              } else {
                convertGenshinWishesExportToWishTally(wishTallySource, cacheConvertedSource,autoImportSheet);
              }
            } else {
              title = "Error";
              message = "Must provide Wish Tally sheet URL, check cell "+RANGE_WISH_TALLY_URL_USER_INPUT;
            }
          } else {
            title = "Error";
            if (fileTypeSelection == RANGE_EXPORT_GENSHIN_WISHES_FILE_TYPE) {
              message = "Source is not an CSV file, check cell "+RANGE_WISH_TALLY_URL_USER_INPUT;
            } else {
              message = "Source is not an Microsoft Excel file, check cell "+RANGE_WISH_TALLY_URL_USER_INPUT;
            }
          }
        } else {
          title = "Error";
          message = "Must provide source file URL to import wishes, check cell "+RANGE_FILE_URL_USER_INPUT;
        }
      } else {
        title = "Error";
        message = "Selected file type is not recognised, check cell "+RANGE_FILE_TYPE_SELECTION;
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
    if (exportFormat == PAIMON_MOE_EXPORT_SHEET_NAME) {
      range.sort([{column: 6, ascending: true}]);
    } else if (exportFormat == GENSHIN_WISHES_EXPORT_SHEET_NAME) {
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
    if (exportFormat == PAIMON_MOE_EXPORT_SHEET_NAME) {
      exportSheet.getRange(3, 9, lastRowWithoutTitle, 6).setBackground("lightgrey");
    } else if (exportFormat == GENSHIN_WISHES_EXPORT_SHEET_NAME) {
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

// paimon.moe
function paimonMoeExportAdjustAndSortSheet() {
  exportAdjustFormat(PAIMON_MOE_EXPORT_SHEET_NAME);
  exportSortSheet(PAIMON_MOE_EXPORT_SHEET_NAME);
}

function paimonMoeExportAdjustFormat() {
  exportAdjustFormat(PAIMON_MOE_EXPORT_SHEET_NAME);
}

function paimonMoeExportSortSheet() {
  exportSortSheet(PAIMON_MOE_EXPORT_SHEET_NAME);
}

function paimonMoeExportClearSheet() {
  exportClearSheet(PAIMON_MOE_EXPORT_SHEET_NAME);
}