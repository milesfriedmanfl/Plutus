package ExcelUtils.Sheets;

import ExcelUtils.ExcelCellLocation;
import ExcelUtils.ExcelColumn;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.List;
import java.util.stream.IntStream;

public class PaycheckAnalysisSheet extends ExcelSheet {
    public PaycheckAnalysisSheet(XSSFWorkbook workbook, String sheetName) {
        super(workbook, sheetName);
    }

    @Override
    protected void setLocationsOfCellsToCarryOver() {
        this.cellsToCarryOver.addAll(PAY_PERIOD_VALUE_LOCATIONS());
        this.cellsToCarryOver.addAll(INCOME_PAY_LINE_ITEM_LOCATIONS());
        this.cellsToCarryOver.addAll(INCOME_PAY_VALUE_LOCATIONS());
        this.cellsToCarryOver.addAll(DEDUCTIONS_BENEFIT_FLAT_RATE_LINE_ITEM_LOCATIONS());
        this.cellsToCarryOver.addAll(DEDUCTIONS_BENEFIT_FLAT_RATE_VALUE_LOCATIONS());
        this.cellsToCarryOver.addAll(DEDUCTIONS_BENEFIT_VARIABLE_RATE_LINE_ITEM_LOCATIONS());
        this.cellsToCarryOver.addAll(DEDUCTIONS_BENEFIT_VARIABLE_RATE_VALUE_LOCATIONS());
        this.cellsToCarryOver.addAll(DEDUCTIONS_TAXES_LINE_ITEM_LOCATIONS());
        this.cellsToCarryOver.addAll(DEDUCTIONS_TAXES_VALUE_LOCATIONS());
    }

    @Override
    protected void setLocationsOfFormulaCellsToReevaluate() {
        this.formulaCellsToReevaluate.addAll(INCOME_ROW_FORMULA_CELL_LOCATIONS());
        this.formulaCellsToReevaluate.addAll(BENEFITS_FLAT_RATES_FORMULA_CELL_LOCATIONS());
        this.formulaCellsToReevaluate.addAll(BENEFITS_VARIABLE_RATES_FORMULA_CELL_LOCATIONS());
        this.formulaCellsToReevaluate.addAll(TAXES_FORMULA_CELL_LOCATIONS());
        this.formulaCellsToReevaluate.addAll(TOTAL_DEDUCTIONS_FORMULA_CELL_LOCATIONS());
        this.formulaCellsToReevaluate.addAll(NET_INCOME_FORMULA_CELL_LOCATIONS());
    }

    private static List<ExcelCellLocation> PAY_PERIOD_VALUE_LOCATIONS() {
        ExcelColumn[] columnsOfCellsToEdit = new ExcelColumn[] {
                ExcelColumn.G, ExcelColumn.H, ExcelColumn.I, ExcelColumn.J, ExcelColumn.K, ExcelColumn.L
        };
        int[] rowsOfCellsToEdit = new int[] {3};
        return createCellListPermutations(columnsOfCellsToEdit, rowsOfCellsToEdit, ExcelCellLocation.VALUE_TYPE.STRING);
    }

    private static List<ExcelCellLocation> INCOME_PAY_LINE_ITEM_LOCATIONS() {
        int[] rowsOfCellsToEdit = IntStream.range(7, 10).toArray();
        return getLineItemCellLocationsForRows(rowsOfCellsToEdit);
    }

    private static List<ExcelCellLocation> INCOME_PAY_VALUE_LOCATIONS() {
        ExcelColumn[] columnsOfCellsToEdit = new ExcelColumn[] {
                ExcelColumn.G, ExcelColumn.H, ExcelColumn.I, ExcelColumn.J, ExcelColumn.K, ExcelColumn.L
        };
        int[] rowsOfCellsToEdit = IntStream.range(7, 10).toArray();
        return createCellListPermutations(columnsOfCellsToEdit, rowsOfCellsToEdit, ExcelCellLocation.VALUE_TYPE.NUMERIC);
    }

    private static List<ExcelCellLocation> DEDUCTIONS_BENEFIT_FLAT_RATE_LINE_ITEM_LOCATIONS() {
        int[] rowsOfCellsToEdit = IntStream.range(15, 24).toArray();
        return getLineItemCellLocationsForRows(rowsOfCellsToEdit);
    }

    private static List<ExcelCellLocation> DEDUCTIONS_BENEFIT_FLAT_RATE_VALUE_LOCATIONS() {
        ExcelColumn[] columnsOfCellsToEdit = new ExcelColumn[] {
                ExcelColumn.G, ExcelColumn.H, ExcelColumn.I, ExcelColumn.J, ExcelColumn.K, ExcelColumn.L
        };
        int[] rowsOfCellsToEdit = IntStream.range(15, 24).toArray();
        return createCellListPermutations(columnsOfCellsToEdit, rowsOfCellsToEdit, ExcelCellLocation.VALUE_TYPE.NUMERIC);
    }

    private static List<ExcelCellLocation> DEDUCTIONS_BENEFIT_VARIABLE_RATE_LINE_ITEM_LOCATIONS() {
        int[] rowsOfCellsToEdit = IntStream.range(27, 31).toArray();
        return getLineItemCellLocationsForRows(rowsOfCellsToEdit);
    }

    private static List<ExcelCellLocation> DEDUCTIONS_BENEFIT_VARIABLE_RATE_VALUE_LOCATIONS() {
        ExcelColumn[] columnsOfCellsToEdit = new ExcelColumn[] {ExcelColumn.F};
        int[] rowsOfCellsToEdit = IntStream.range(27, 31).toArray();
        return createCellListPermutations(columnsOfCellsToEdit, rowsOfCellsToEdit, ExcelCellLocation.VALUE_TYPE.NUMERIC);
    }

    private static List<ExcelCellLocation> DEDUCTIONS_TAXES_LINE_ITEM_LOCATIONS() {
        int[] rowsOfCellsToEdit = IntStream.range(34, 40).toArray();
        return getLineItemCellLocationsForRows(rowsOfCellsToEdit);
    }

    private static List<ExcelCellLocation> DEDUCTIONS_TAXES_VALUE_LOCATIONS() {
        ExcelColumn[] columnsOfCellsToEdit = new ExcelColumn[] {ExcelColumn.F};
        int[] rowsOfCellsToEdit = IntStream.range(34, 40).toArray();
        return createCellListPermutations(columnsOfCellsToEdit, rowsOfCellsToEdit, ExcelCellLocation.VALUE_TYPE.NUMERIC);
    }

    private static List<ExcelCellLocation> INCOME_ROW_FORMULA_CELL_LOCATIONS() {
        int[] rowsOfCellsToEdit = IntStream.range(7, 11).toArray();
        return getFormulaCellLocationsForRows(rowsOfCellsToEdit);
    }

    private static List<ExcelCellLocation> BENEFITS_FLAT_RATES_FORMULA_CELL_LOCATIONS() {
        int[] rowsOfCellsToEdit = IntStream.range(15, 25).toArray();
        return getFormulaCellLocationsForRows(rowsOfCellsToEdit);
    }

    private static List<ExcelCellLocation> BENEFITS_VARIABLE_RATES_FORMULA_CELL_LOCATIONS() {
        int[] rowsOfCellsToEdit = IntStream.range(27, 32).toArray();
        return getFormulaCellLocationsForRows(rowsOfCellsToEdit);
    }

    private static List<ExcelCellLocation> TAXES_FORMULA_CELL_LOCATIONS() {
        int[] rowsOfCellsToEdit = IntStream.range(34, 41).toArray();
        return getFormulaCellLocationsForRows(rowsOfCellsToEdit);
    }

    private static List<ExcelCellLocation> TOTAL_DEDUCTIONS_FORMULA_CELL_LOCATIONS() {
        int[] rowsOfCellsToEdit = new int[]{42};
        return getFormulaCellLocationsForRows(rowsOfCellsToEdit);
    }

    private static List<ExcelCellLocation> NET_INCOME_FORMULA_CELL_LOCATIONS() {
        int[] rowsOfCellsToEdit = new int[]{45};
        return getFormulaCellLocationsForRows(rowsOfCellsToEdit);
    }

    private static List<ExcelCellLocation> getLineItemCellLocationsForRows(int[] rows) {
        ExcelColumn[] columnsOfCellsToEdit = new ExcelColumn[] {ExcelColumn.C};
        return createCellListPermutations(columnsOfCellsToEdit, rows, ExcelCellLocation.VALUE_TYPE.STRING);
    }

    private static List<ExcelCellLocation> getFormulaCellLocationsForRows(int[] rows) {
        ExcelColumn[] columnsOfCellsToEdit = new ExcelColumn[] {
                // CALCULATION COLUMN NUMS
                ExcelColumn.M, ExcelColumn.N, ExcelColumn.O
        };
        return createCellListPermutations(columnsOfCellsToEdit, rows, ExcelCellLocation.VALUE_TYPE.NUMERIC);
    }
}
