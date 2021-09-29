/*
 * Wish Tally - Conversion
 * Version 1.5 made by yippy
 * https://github.com/Yippy/wish-tally-conversion-sheet
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
  .addSubMenu(ui.createMenu('genshin-wishes')
             .addItem('Clear', 'genshinWishesExportClearSheet')
             .addItem('Adjust Format', 'genshinWishesExportAdjustFormat')
             .addItem('Sort Wish Count', 'genshinWishesExportSortSheet')
             .addItem('Adjust and Sort', 'genshinWishesExportAdjustAndSortSheet'))
  .addSeparator()
  .addItem('Auto Import', 'autoImportToWishTally')
  .addSeparator()
  .addToUi();
}