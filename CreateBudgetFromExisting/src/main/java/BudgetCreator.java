import ExcelUtils.ExcelCellLocation;
import ExcelUtils.ExcelWorkbook;
import ExcelUtils.Sheets.ExcelSheet;
import ExcelUtils.Sheets.SavingsAnalysisSheet;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;

import java.io.IOException;
import java.util.List;
import java.util.Optional;

public class BudgetCreator {
    private static final String PAYCHECK_SHEET_NAME = "Paycheck Analysis";
    private static final String BUDGET_SHEET_NAME = "Budget Analysis";
    private static final String SAVINGS_SHEET_NAME = "Savings Analysis";
    private enum ARGS {
        EXISTING_BUDGET_PATH(0), // Arg 0
        NEW_BUDGET_PATH(1); // Arg 1

        private final int value;
        private ARGS(int value) {
            this.value = value;
        }

        public int getValue() {
            return this.value;
        }
    }

    public static void main(String[] args) throws IOException {
        if (args.length == 0) {
            System.out.println("Could not execute program due to lack of valid arguments.");
            return;
        }

        try {
            String[] budgetTemplateSheetNames = new String[]{PAYCHECK_SHEET_NAME, BUDGET_SHEET_NAME, SAVINGS_SHEET_NAME};
            ExcelWorkbook existingBudget = new ExcelWorkbook(args[ARGS.EXISTING_BUDGET_PATH.getValue()], budgetTemplateSheetNames);
            ExcelWorkbook newBudget = new ExcelWorkbook(args[ARGS.NEW_BUDGET_PATH.getValue()], budgetTemplateSheetNames);

            // Edits newBudget workbook
            for (String sheetName: budgetTemplateSheetNames) {
                ExcelSheet existingBudgetCurrentSheet = existingBudget.getSheet(sheetName);
                ExcelSheet newBudgetCurrentSheet = newBudget.getSheet(sheetName);

                updateSheetCellValuesForNewBudget(existingBudgetCurrentSheet, newBudgetCurrentSheet);
                reRunSheetFormulas(newBudget, sheetName);
            }

            // Save the changes made to the newBudget workbook
            newBudget.saveEditsToWorkbook();
        } catch (IOException e) {
            System.out.println("Error creating new budget file. Please make sure that the file names you have specified are correct and try again.");
            throw new IOException();
        }
    }

    /**
     * Edits cells of the sheet for newBudget with the proper carry over and overwrite values from the same sheet
     * in the existingBudget.
     */
    private static void updateSheetCellValuesForNewBudget(ExcelSheet existingBudgetSheet, ExcelSheet newBudgetSheet) {
        List<ExcelCellLocation> carryOverCellLocations = existingBudgetSheet.getCellsToCarryOver();
        for (int i = 0; i < carryOverCellLocations.size(); i++) {
            ExcelCellLocation locationOfCellToCarryOver = carryOverCellLocations.get(i);
            XSSFCell cellToCarryOver = existingBudgetSheet.getCell(locationOfCellToCarryOver);
            XSSFCellStyle carryOverCellStyle = cellToCarryOver.getCellStyle();
            StylesTable carryOverCellStyleSource = cellToCarryOver.getSheet().getWorkbook().getStylesSource();
            RichTextString carryOverCellComment =
                    Optional.ofNullable(cellToCarryOver)
                    .map(XSSFCell::getCellComment)
                    .map(Comment::getString)
                            .orElse(null);

            if (SAVINGS_SHEET_NAME.equals(newBudgetSheet.getSheetName())) {
                // Overwrite cell to overwrite contents in new budget using carry over cell contents of the old budget.
                // This is the only sheet in which the cells to carry over are not the same as the cells to overwrite.
                ExcelCellLocation locationOfCellToOverwrite = ((SavingsAnalysisSheet)newBudgetSheet).getCellToOverwrite(i);
                double carryOverCellValue = cellToCarryOver.getNumericCellValue();
                newBudgetSheet.setCellContents(
                        locationOfCellToOverwrite,
                        carryOverCellValue,
                        carryOverCellStyle,
                        carryOverCellStyleSource,
                        carryOverCellComment
                );
            } else {
                // Overwrite carry over cell contents of the new budget using carry over cell contents of the old budget
                if (locationOfCellToCarryOver.getValueType() == ExcelCellLocation.VALUE_TYPE.STRING) {
                    String carryOverCellValue = cellToCarryOver.getStringCellValue();
                    newBudgetSheet.setCellContents(locationOfCellToCarryOver, carryOverCellValue, carryOverCellStyle, carryOverCellStyleSource, carryOverCellComment);
                } else {
                    double carryOverCellValue = cellToCarryOver.getNumericCellValue();
                    newBudgetSheet.setCellContents(
                            locationOfCellToCarryOver,
                            carryOverCellValue,
                            carryOverCellStyle,
                            carryOverCellStyleSource,
                            carryOverCellComment
                    );
                }
            }
        }
    }

    /**
     * Forces the formulas for the given sheet of the passed in workbook to calculate new values. This method should be
     * used after cells of the sheet are edited that the sheet's formulas depend on.
     */
    private static void reRunSheetFormulas(ExcelWorkbook workbook, String sheetName) {
        ExcelSheet sheet = workbook.getSheet(sheetName);
        FormulaEvaluator evaluator = workbook.getWorkbook().getCreationHelper().createFormulaEvaluator();
        for (ExcelCellLocation cellLocation: sheet.getFormulaCellsToReevaluate()) {
            XSSFCell formulaCell = sheet.getCell(cellLocation);
            evaluator.evaluateFormulaCell(formulaCell);
        }
    }
}
