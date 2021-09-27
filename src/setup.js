/*
 * Wish Tally - Conversion
 * Version 1.5 made by yippym
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