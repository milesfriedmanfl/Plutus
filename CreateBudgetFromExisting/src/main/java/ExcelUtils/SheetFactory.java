package ExcelUtils;

import ExcelUtils.Sheets.BudgetAnalysisSheet;
import ExcelUtils.Sheets.ExcelSheet;
import ExcelUtils.Sheets.PaycheckAnalysisSheet;
import ExcelUtils.Sheets.SavingsAnalysisSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class SheetFactory {
    public static ExcelSheet getSheet(XSSFWorkbook workbook, String sheet) {
        String sheetWithoutPrefix = sheet.toUpperCase().substring(0, sheet.toUpperCase().lastIndexOf("ANALYSIS")).trim();
        if ("PAYCHECK".equals(sheetWithoutPrefix)) {
            return new PaycheckAnalysisSheet(workbook, sheet);
        } else if ("BUDGET".equals(sheetWithoutPrefix)) {
            return new BudgetAnalysisSheet(workbook, sheet);
        } else if ("SAVINGS".equals(sheetWithoutPrefix)) {
            return new SavingsAnalysisSheet(workbook, sheet);
        } else {
            System.out.println("Error creating budget, could not find sheet in factory with name: " + sheet + ".");
            return null;
        }
    }
}
