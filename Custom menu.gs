function onOpen() {
  ui = SpreadsheetApp.getUi().createMenu('Custom menu')
        .addItem('Update NHMIS Reporting Rates', 'updateNHMISReportingRates')
        .addItem('Update codes', 'update_codes_in_sheet')
        .addItem('Update extract info in sheet', 'update_extract_info_in_sheet')
        .addItem('Manually update sheet with weekly or daily data', 'updateDataWeekly')
        .addToUi();
}
