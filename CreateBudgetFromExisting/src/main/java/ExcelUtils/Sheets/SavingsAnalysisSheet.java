package ExcelUtils.Sheets;

import ExcelUtils.ExcelCellLocation;
import ExcelUtils.ExcelColumn;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.ArrayList;
import java.util.List;
import java.util.stream.IntStream;

public class SavingsAnalysisSheet extends ExcelSheet {
    // For this sheet only the cellsToOverwrite differs from the carryOverCells, so we need another list for this.
    private List<ExcelCellLocation> cellsToOverwrite = new ArrayList<>();

    // It is however important that the rows for the carryOverCells and cellsToOverwrite line up, so we also need to use
    // the same lists of rows when setting up both cell lists.
    private static final int[] RESTRICTED_SAVINGS_CELL_ROWS = IntStream.range(67, 77).toArray();
    private static final int[] SUPERFLUOUS_SAVINGS_CELL_ROWS = IntStream.range(80, 86).toArray();

    public SavingsAnalysisSheet(XSSFWorkbook workbook, String sheetName) {
        super(workbook, sheetName);
        this.setCellsToOverwrite();
    }

    public ExcelCellLocation getCellToOverwrite(int index) {
        return this.cellsToOverwrite.get(index);
    }

    @Override
    protected void setLocationsOfCellsToCarryOver() {
        this.cellsToCarryOver.addAll(EXPECTED_RESTRICTED_SAVINGS_CARRY_OVER_VALUE_LOCATIONS());
        this.cellsToCarryOver.addAll(EXPECTED_SUPERFLUOUS_SAVINGS_CARRY_OVER_VALUE_LOCATIONS());
    }

    @Override
    protected void setLocationsOfFormulaCellsToReevaluate() {
        this.formulaCellsToReevaluate.addAll(INFLOWS_RESTRICTED_SAVINGS_FORMULA_LOCATIONS());
        this.formulaCellsToReevaluate.addAll(INFLOWS_SUPERFLUOUS_SAVINGS_FORMULA_LOCATIONS());
        this.formulaCellsToReevaluate.addAll(INFLOWS_TOTAL_TRANSFERRED_IN_FORMULA_LOCATIONS());
        this.formulaCellsToReevaluate.addAll(OUTFLOWS_RESTRICTED_SAVINGS_FORMULA_LOCATIONS());
        this.formulaCellsToReevaluate.addAll(EXPECTED_RESTRICTED_FORMULA_LOCATIONS());
        this.formulaCellsToReevaluate.addAll(OUTFLOWS_TOTAL_TRANSFERRED_OUT_FORMULA_LOCATIONS());
        this.formulaCellsToReevaluate.addAll(NET_SAVED_FORMULA_LOCATIONS());
        this.formulaCellsToReevaluate.addAll(EXPECTED_RESTRICTED_ACCOUNT_FORMULA_LOCATIONS());
        this.formulaCellsToReevaluate.addAll(EXPECTED_SUPERFLUOUS_ACCOUNT_FORMULA_LOCATIONS());
    }

    private void setCellsToOverwrite() {
        this.cellsToOverwrite.addAll(EXPECTED_RESTRICTED_SAVINGS_OVERWRITE_VALUE_LOCATIONS());
        this.cellsToOverwrite.addAll(EXPECTED_SUPERFLUOUS_SAVINGS_OVERWRITE_VALUE_LOCATIONS());
    }

    private static List<ExcelCellLocation> EXPECTED_RESTRICTED_SAVINGS_CARRY_OVER_VALUE_LOCATIONS() {
        return getCarryOverValueCellLocationsForRows(RESTRICTED_SAVINGS_CELL_ROWS);
    }

    private static List<ExcelCellLocation> EXPECTED_RESTRICTED_SAVINGS_OVERWRITE_VALUE_LOCATIONS() {
        return getOverwriteValueCellLocationsForRows(RESTRICTED_SAVINGS_CELL_ROWS);
    }

    private static List<ExcelCellLocation> EXPECTED_SUPERFLUOUS_SAVINGS_CARRY_OVER_VALUE_LOCATIONS() {
        return getCarryOverValueCellLocationsForRows(SUPERFLUOUS_SAVINGS_CELL_ROWS);
    }

    private static List<ExcelCellLocation> EXPECTED_SUPERFLUOUS_SAVINGS_OVERWRITE_VALUE_LOCATIONS() {
        return getOverwriteValueCellLocationsForRows(SUPERFLUOUS_SAVINGS_CELL_ROWS);
    }

    private static List<ExcelCellLocation> INFLOWS_RESTRICTED_SAVINGS_FORMULA_LOCATIONS() {
        int[] rowsOfCellsToEdit = IntStream.range(13, 24).toArray();
        return getFormulaCellLocationsForRows(rowsOfCellsToEdit);
    }

    private static List<ExcelCellLocation> INFLOWS_SUPERFLUOUS_SAVINGS_FORMULA_LOCATIONS() {
        int[] rowsOfCellsToEdit = IntStream.range(27, 33).toArray();
        return getFormulaCellLocationsForRows(rowsOfCellsToEdit);
    }

    private static List<ExcelCellLocation> INFLOWS_TOTAL_TRANSFERRED_IN_FORMULA_LOCATIONS() {
        int[] rowsOfCellsToEdit = new int[]{34};
        return getFormulaCellLocationsForRows(rowsOfCellsToEdit);
    }

    private static List<ExcelCellLocation> OUTFLOWS_RESTRICTED_SAVINGS_FORMULA_LOCATIONS() {
        int[] rowsOfCellsToEdit = IntStream.range(39, 50).toArray();
        return getFormulaCellLocationsForRows(rowsOfCellsToEdit);
    }

    private static List<ExcelCellLocation> EXPECTED_RESTRICTED_FORMULA_LOCATIONS() {
        int[] rowsOfCellsToEdit = IntStream.range(53, 59).toArray();
        return getFormulaCellLocationsForRows(rowsOfCellsToEdit);
    }

    private static List<ExcelCellLocation> OUTFLOWS_TOTAL_TRANSFERRED_OUT_FORMULA_LOCATIONS() {
        int[] rowsOfCellsToEdit = new int[]{60};
        return getFormulaCellLocationsForRows(rowsOfCellsToEdit);
    }

    private static List<ExcelCellLocation> NET_SAVED_FORMULA_LOCATIONS() {
        int[] rowsOfCellsToEdit = new int[]{62};
        return getFormulaCellLocationsForRows(rowsOfCellsToEdit);
    }

    private static List<ExcelCellLocation> EXPECTED_RESTRICTED_ACCOUNT_FORMULA_LOCATIONS() {
        int[] rowsOfCellsToEdit = IntStream.range(67, 78).toArray();
        return getFormulaCellLocationsForRows(rowsOfCellsToEdit);
    }

    private static List<ExcelCellLocation> EXPECTED_SUPERFLUOUS_ACCOUNT_FORMULA_LOCATIONS() {
        int[] rowsOfCellsToEdit = IntStream.range(80, 87).toArray();
        return getFormulaCellLocationsForRows(rowsOfCellsToEdit);
    }

    private static List<ExcelCellLocation> getCarryOverValueCellLocationsForRows(int[] rowsOfCellsToEdit) {
        ExcelColumn[] columnsOfCellsToEdit = new ExcelColumn[] {ExcelColumn.S}; // Dec Expected Values Column
        return createCellListPermutations(columnsOfCellsToEdit, rowsOfCellsToEdit, ExcelCellLocation.VALUE_TYPE.NUMERIC);
    }

    private static List<ExcelCellLocation> getOverwriteValueCellLocationsForRows(int[] rowsOfCellsToEdit) {
        ExcelColumn[] columnsOfCellsToEdit = new ExcelColumn[] {ExcelColumn.G}; // Start Values Column
        return createCellListPermutations(columnsOfCellsToEdit, rowsOfCellsToEdit, ExcelCellLocation.VALUE_TYPE.NUMERIC);
    }

    private static List<ExcelCellLocation> getFormulaCellLocationsForRows(int[] rowsOfCellsToEdit) {
        // MONTHLY VALUE COLUMN NUMS
        ExcelColumn[] columnsOfCellsToEdit = new ExcelColumn[] {
                ExcelColumn.C, ExcelColumn.G, ExcelColumn.H, ExcelColumn.I, ExcelColumn.J, ExcelColumn.K, ExcelColumn.L, ExcelColumn.M,
                ExcelColumn.N, ExcelColumn.O, ExcelColumn.P, ExcelColumn.Q, ExcelColumn.R, ExcelColumn.S
        };
        return createCellListPermutations(columnsOfCellsToEdit, rowsOfCellsToEdit, ExcelCellLocation.VALUE_TYPE.NUMERIC);
    }
}
