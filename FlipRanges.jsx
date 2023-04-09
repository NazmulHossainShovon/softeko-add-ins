import React, { useEffect, useState } from "react";
import Title from "../../../../shared/reusableComponents/Title";
import HorizontalRadioButton from "../../../../shared/reusableComponents/HorizontalRadioButton";
import SourceInputDataLogic from "../../../../shared/data/SourceInputDataLogic";
import { WarningModal } from "../../../../shared/reusableComponents/WarningModal";
import ConfirmationDialogue from "../../../../shared/reusableComponents/ConfirmationDialogue";
import ButtonConfirmationDialogue from "../../../../shared/reusableComponents/ButtonConfirmationDialogue";
import UserGuide from "./elements/UserGuide";
import { Checkbox, FormControlLabel, FormGroup, Paper } from "@mui/material";
import ExcelHelper from "./excelHelper";

const radioInfo = [
  { id: "1", value: "horizontally", label: "Horizontally" },
  { id: "2", value: "vertically", label: "Vertically" },
];

export default function FlipRanges({ isOfficeInitialized }) {
  const [ranges, setRanges] = React.useState(" ");
  const [selection, setSelection] = React.useState("horizontally");
  let [rowNo, setRowNo] = React.useState("");
  let [colNo, setColNo] = React.useState("");
  const [rowIndex, setRowIndex] = React.useState("");
  const [columnIndex, setColumnIndex] = React.useState("");
  const [sourceValues, setSourceValues] = React.useState("");
  const [backup, setBackup] = React.useState(false);
  const [render, setRender] = React.useState(false);
  const [adjustCellRefs, setAdjustCellRefs] = useState(false);
  const [keepFormatting, setKeepFormatting] = useState(false);

  const [warningOpen, setWarningOpen] = React.useState(false);
  const handleWarningOpen = () => setWarningOpen(true);
  const handleWarningClose = () => setWarningOpen(false);

  const transferData = (data) => {
    setRowNo(data.rowNo);
    setColNo(data.colNo);
    setSourceValues(data.sourceValues);
    setRowIndex(data.rowIndex);
    setColumnIndex(data.columnIndex);
  };
  const transferSourceRange = (data) => {
    setRanges(data.sourceRanges);
  };

  const selectionChangeHandler = (e) => {
    setSelection(e.target.value);
  };

  const transferBackupInfo = (data) => {
    setBackup(data.backup);
  };
  const backupWorksheet = async () => {
    try {
      await Excel.run(async (context) => {
        let myWorkbook = context.workbook;
        let sampleSheet = myWorkbook.worksheets.getActiveWorksheet();
        let copiedSheet = sampleSheet.copy(Excel.WorksheetPositionType.after, sampleSheet);
        await context.sync();
      });
    } catch (error) {
      console.log(error);
    }
  };

  const horizontalFlip = async () => {
    try {
      // eslint-disable-next-line no-undef
      await Excel.run(async (context) => {
        let sheet;
        let backup_ranges;
        if (backup) {
          let activeSheet = context.workbook.worksheets.getActiveWorksheet();
          sheet = activeSheet.copy(Excel.WorksheetPositionType.after, activeSheet);
          await context.sync();
          backup_ranges = ranges.split("!")[1];
        } else {
          backup_ranges = ranges;
          sheet = context.workbook.worksheets.getActiveWorksheet();
        }
        const range = sheet.getRange(backup_ranges);
        if (adjustCellRefs) {
          colNo = colNo - 1;
        }

        //flipping upper half of the rows
        for (let i = 0; i < parseInt(rowNo / 2); i++) {
          for (let j = 0; j < colNo; j++) {
            range.getCell(i, j).values = sourceValues[rowNo - 1 - i][j];
          }
        }
        //flipping lower half of the rows
        for (let i = 0; i < parseInt(rowNo / 2); i++) {
          for (let j = 0; j < colNo; j++) {
            range.getCell(rowNo - 1 - i, j).values = sourceValues[i][j];
          }
        }
        await context.sync();
        handleWarningClose();
        setRender(!render);
      });
    } catch (error) {
      console.error(error);
    }
  };

  const verticalFlip = async () => {
    try {
      await Excel.run(async (context) => {
        let sheet;
        let backup_ranges;
        if (backup) {
          let activeSheet = context.workbook.worksheets.getActiveWorksheet();
          sheet = activeSheet.copy(Excel.WorksheetPositionType.after, activeSheet);
          await context.sync();
          backup_ranges = ranges.split("!")[1];
        } else {
          backup_ranges = ranges;
          sheet = context.workbook.worksheets.getActiveWorksheet();
        }
        const range = sheet.getRange(backup_ranges);

        if (adjustCellRefs) {
          rowNo = rowNo - 1;
        }
        //flipping left half of the columns
        for (let i = 0; i < parseInt(colNo / 2); i++) {
          for (let j = 0; j < rowNo; j++) {
            range.getCell(j, i).values = sourceValues[j][colNo - (i + 1)];
          }
        }
        //flipping right half of the columns
        for (let i = 0; i < parseInt(colNo / 2); i++) {
          for (let j = 0; j < rowNo; j++) {
            range.getCell(j, colNo - (i + 1)).values = sourceValues[j][i];
          }
        }
        {
          keepFormatting && formatKeeper(range, context);
        }
        handleWarningClose();
        setRender(!render);
      });
    } catch (error) {
      console.error(error);
    }
  };

  const formatKeeper = async (dataRange, context) => {
    let sourceRanges = [];
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const dummyAddress = await getDummyAddress(sheet, dataRange, context);
    const dummyRange = sheet.getRange(dummyAddress);
    let dummySourceRange = [];

    dummyRange.copyFrom(dataRange, Excel.RangeCopyType.formats);

    for (let i = 0; i < colNo; i++) {
      console.log("hello");
      const singleCell = dataRange.getCell(0, i);

      const dummySingleCell = dummyRange.getCell(0, i);
      singleCell.load("address");
      dummySingleCell.load("address");
      await context.sync();

      sourceRanges.push(singleCell.address);
      dummySourceRange.push(dummySingleCell.address);
    }

    const targetRanges = reverseArray(sourceRanges);
    for (let j = 0; j < sourceValues[0].length; j++) {
      ExcelHelper.copyRangeFormat(dummySourceRange[j], targetRanges[j]);
      await context.sync();
    }

    function reverseArray(arr) {
      const reversedArr = [];
      for (let i = arr.length - 1; i >= 0; i--) {
        reversedArr.push(arr[i]);
      }
      return reversedArr;
    }

    dummyRange.clear(Excel.ClearApplyTo.all);
  };

  const getDummyAddress = async (sheet, range, context) => {
    const startCell = sheet.getRange("IPV850000");
    startCell.load(["rowIndex", "columnIndex"]);
    await context.sync();

    const startRowIndex = startCell.rowIndex;
    const startColumnIndex = startCell.columnIndex;

    range.load(["rowCount", "columnCount"]);
    await context.sync();

    const numRows = range.rowCount;
    const numColumns = range.columnCount;

    const outputRange = sheet.getRangeByIndexes(startRowIndex, startColumnIndex, numRows, numColumns);
    outputRange.load("address");
    await context.sync();

    return outputRange.address;
  };

  return (
    <React.Fragment>
      <Title title="Flip Ranges" userGuide={<UserGuide />} />
      <SourceInputDataLogic
        isOfficeInitialized={isOfficeInitialized}
        transferData={transferData}
        transferSourceRange={transferSourceRange}
        transferBackupInfo={transferBackupInfo}
        label="Source Range"
        alertName="Source Range Empty Error"
        alertMsg="Please! Select or Type a Range"
        render={render}
      >
        <div className="centered">
          {selection === "horizontally" && (
            <img
              src="https://milleary.sirv.com/Images/flip_horizonatally.png"
              width="262"
              height="128"
              alt="horizontally"
            />
          )}
          {selection === "vertically" && (
            <img src="https://milleary.sirv.com/Images/flip_vertically.png" width="289" height="128" alt="vertically" />
          )}
        </div>

        <HorizontalRadioButton
          title="Selection Type"
          defaultValue="horizontally"
          formData={radioInfo}
          onChange={selectionChangeHandler}
        />
        <Paper elevation={1} sx={{ marginBottom: "10px", marginTop: "10px", padding: "5px" }}>
          <FormGroup>
            <FormControlLabel
              sx={{
                span: { padding: "3px" },
                "& .MuiTypography-root": { fontSize: ".85rem", fontWeight: "500" },
                svg: { width: ".95rem", height: ".95rem", marginLeft: "10px" },
              }}
              control={<Checkbox onClick={() => setAdjustCellRefs(!adjustCellRefs)} />}
              label="Adjust Cell References"
            />
            <FormControlLabel
              sx={{
                span: { padding: "3px" },
                "& .MuiTypography-root": { fontSize: ".85rem", fontWeight: "500" },
                svg: { width: ".95rem", height: ".95rem", marginLeft: "10px" },
              }}
              control={<Checkbox onClick={() => setKeepFormatting(!keepFormatting)} />}
              label="Keep Formatting"
            />
          </FormGroup>
        </Paper>
      </SourceInputDataLogic>
      <WarningModal open={warningOpen} onClose={handleWarningClose}>
        {selection === "horizontally" && <ConfirmationDialogue onClick={horizontalFlip} deny={handleWarningClose} />}
        {selection === "vertically" && <ConfirmationDialogue onClick={verticalFlip} deny={handleWarningClose} />}
      </WarningModal>
      <ButtonConfirmationDialogue size="sm" onClick={handleWarningOpen} selectedRange={ranges} targetRange="nothing" />
    </React.Fragment>
  );
}
