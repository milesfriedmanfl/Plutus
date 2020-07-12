package ExcelUtils.Sheets;

import ExcelUtils.ExcelCellLocation;
import ExcelUtils.ExcelColumn;
import org.apache.poi.xssf.usermodel.*;

import java.util.ArrayList;
import java.util.List;

public abstract class ExcelSheet {
    protected String sheetName;
    protected XSSFSheet sheet;
    protected List<ExcelCellLocation> cellsToCarryOver = new ArrayList<>();
    protected List<ExcelCellLocation> formulaCellsToReevaluate = new ArrayList<>();

    protected ExcelSheet(XSSFWorkbook workbook, String sheetName) {
        this.sheetName = sheetName;
        this.sheet = workbook.getSheet(sheetName);
        this.setLocationsOfCellsToCarryOver();
        this.setLocationsOfFormulaCellsToReevaluate();
    }

    protected static List<ExcelCellLocation> createCellListPermutations(ExcelColumn[] cols, int[] rows, ExcelCellLocation.VALUE_TYPE valueType) {
        List<ExcelCellLocation> result = new ArrayList<>();
        for (ExcelColumn col: cols) {
            for (int row: rows) {
                result.add(new ExcelCellLocation(col, row, valueType));
            }
        }
        return result;
    }

    protected abstract void setLocationsOfCellsToCarryOver();
    
    protected abstract void setLocationsOfFormulaCellsToReevaluate();

    public XSSFCell getCell(ExcelCellLocation cell) {
        XSSFRow sheetRow = this.sheet.getRow(cell.getRow());
        return sheetRow.getCell(cell.getColumn());
    }

    public void setCellContents(ExcelCellLocation cell, String value, XSSFCellStyle style) {
        XSSFRow sheetRow = this.sheet.getRow(cell.getRow());
        sheetRow.createCell(cell.getColumn()).setCellValue(value);
//        sheetRow.createCell(cell.getColumn()).setCellStyle(style);
    }

    public void setCellContents(ExcelCellLocation cell, double value, XSSFCellStyle style) {
        if (value != 0) {
            XSSFRow sheetRow = this.sheet.getRow(cell.getRow());
            sheetRow.createCell(cell.getColumn()).setCellValue(value);
//        sheetRow.createCell(cell.getColumn()).setCellStyle(style);
        }
    }

    public List<ExcelCellLocation> getCellsToCarryOver() {
        return this.cellsToCarryOver;
    }

    public List<ExcelCellLocation> getFormulaCellsToReevaluate() {
        return this.formulaCellsToReevaluate;
    }

    public String getSheetName() {
        return this.sheetName;
    }
}
