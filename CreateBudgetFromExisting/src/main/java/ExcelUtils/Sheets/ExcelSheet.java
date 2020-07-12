package ExcelUtils.Sheets;

import ExcelUtils.ExcelCellLocation;
import ExcelUtils.ExcelColumn;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellFill;

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

    public List<ExcelCellLocation> getCellsToCarryOver() {
        return this.cellsToCarryOver;
    }

    public List<ExcelCellLocation> getFormulaCellsToReevaluate() {
        return this.formulaCellsToReevaluate;
    }

    public String getSheetName() {
        return this.sheetName;
    }

    public XSSFCell getCell(ExcelCellLocation cell) {
        XSSFRow sheetRow = this.sheet.getRow(cell.getRow());
        return sheetRow.getCell(cell.getColumn());
    }

    public void setCellContents(ExcelCellLocation cellLocation, String value, XSSFCellStyle style, StylesTable stylesSource, RichTextString comment) {
        XSSFRow sheetRow = this.sheet.getRow(cellLocation.getRow());
        XSSFCell cell = sheetRow.createCell(cellLocation.getColumn());
        cell.setCellValue(value);
        this.setCellStyle(cell, style, stylesSource);
        this.setCellComment(cell, comment);
    }

    public void setCellContents(ExcelCellLocation cellLocation, double value, XSSFCellStyle style, StylesTable styleSource, RichTextString comment) {
        if (value != 0) {
            XSSFRow sheetRow = this.sheet.getRow(cellLocation.getRow());
            XSSFCell cell = sheetRow.createCell(cellLocation.getColumn());
            cell.setCellValue(value);
            this.setCellStyle(cell, style, styleSource);
            this.setCellComment(cell, comment);
        }
    }

    /**
     * Due to a bug in apache poi, the existing cell's style source must be used in order to copy fill and border styles
     * @param cell
     * @param existingCellStyle
     * @param existingCellStyleSource
     */
    private void setCellStyle(XSSFCell cell, XSSFCellStyle existingCellStyle, StylesTable existingCellStyleSource) {
        XSSFCellStyle newCellStyle = cell.getSheet().getWorkbook().createCellStyle();
        newCellStyle.cloneStyleFrom(existingCellStyle);
        StylesTable newStyleSource = cell.getSheet().getWorkbook().getStylesSource();

        // Copy over fills
        for (XSSFCellFill fill: existingCellStyleSource.getFills()) {
            XSSFCellFill newFill = new XSSFCellFill(fill.getCTFill(), existingCellStyleSource.getIndexedColors());
            newStyleSource.putFill(newFill);
        }

        // Copy over borders
        for (XSSFCellBorder border: existingCellStyleSource.getBorders()) {
            XSSFCellBorder newBorder = new XSSFCellBorder(border.getCTBorder());
            newStyleSource.putBorder(newBorder);
        }

        cell.setCellStyle(newCellStyle);
    }

    private void setCellComment(XSSFCell cell, RichTextString existingComment) {
        if (existingComment == null) {
            return;
        }

        Drawing drawing = cell.getSheet().createDrawingPatriarch();
        ClientAnchor anchor = cell.getSheet().getWorkbook().getCreationHelper().createClientAnchor();

        // Comment will have a default size of
        int COMMENT_HEIGHT = 10;
        int COMMENT_WIDTH = 30;
        anchor.setCol1(cell.getColumnIndex());
        anchor.setCol2(cell.getColumnIndex() + COMMENT_WIDTH);
        anchor.setRow1(cell.getRowIndex());
        anchor.setRow2(cell.getRowIndex() + COMMENT_HEIGHT);

        // Create comment
        Comment comment = drawing.createCellComment(anchor);
        comment.setString(existingComment);
        cell.setCellComment(comment);
    }
}
