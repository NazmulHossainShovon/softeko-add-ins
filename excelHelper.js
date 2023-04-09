// excelHelper.js

const ExcelHelper = {
  // Existing methods...

  // Method to copy cell formats
  copyRangeFormat: async function (sourceRangeAddress, targetRangeAddress) {
    await Excel.run(async (context) => {
      const activeSheet = context.workbook.worksheets.getActiveWorksheet();
      const sourceRange = activeSheet.getRange(sourceRangeAddress);
      const targetRange = activeSheet.getRange(targetRangeAddress);

      targetRange.copyFrom(sourceRange, Excel.RangeCopyType.formats);

      await context.sync();
    });
  },

  // Existing methods...
};

export default ExcelHelper;
