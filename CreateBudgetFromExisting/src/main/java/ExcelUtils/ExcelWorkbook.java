package ExcelUtils;

import ExcelUtils.Sheets.ExcelSheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

public class ExcelWorkbook {
    private String workbookPath;
    private XSSFWorkbook workbook;
    private Map<String, ExcelSheet> workbookSheets = new HashMap<String, ExcelSheet>();

    public ExcelWorkbook(String filePath, String[] sheetNames) throws IOException {
        if (!filePath.endsWith(".xlsx")) {
            System.out.println("Invalid file: \"" + filePath + "\" provided to ExcelReader. Only files with an .xlsx suffix are accepted.");
        }

        FileInputStream file = new FileInputStream(filePath);
        this.workbookPath = filePath;
        this.workbook = new XSSFWorkbook(file);
        for (String sheetName: sheetNames) {
            this.workbookSheets.put(sheetName,  SheetFactory.getSheet(this.workbook, sheetName));
        }
    }

    public ExcelSheet getSheet(String sheetName) {
        return this.workbookSheets.get(sheetName);
    }

    public Workbook getWorkbook() {
        return this.workbook;
    }

    public void saveEditsToWorkbook() throws IOException {
        FileOutputStream file = new FileOutputStream(this.workbookPath);
        this.workbook.write(file);
        this.workbook.close();
    }
}
