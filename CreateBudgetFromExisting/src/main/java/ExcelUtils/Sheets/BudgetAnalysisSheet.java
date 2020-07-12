package ExcelUtils.Sheets;

import ExcelUtils.ExcelCellLocation;
import ExcelUtils.ExcelColumn;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.List;
import java.util.stream.IntStream;

public class BudgetAnalysisSheet extends ExcelSheet {
    public BudgetAnalysisSheet(XSSFWorkbook workbook, String sheetName) {
        super(workbook, sheetName);
    }

    @Override
    protected void setLocationsOfCellsToCarryOver() {
        this.cellsToCarryOver.addAll(INFLOWS_ADDITIONAL_RECURRING_INCOME_LINE_ITEM_LOCATIONS());
        this.cellsToCarryOver.addAll(INFLOWS_ADDITIONAL_RECURRING_INCOME_VALUE_LOCATIONS());
        this.cellsToCarryOver.addAll(OUTFLOWS_RECURRING_BILLS_LINE_ITEM_LOCATIONS());
        this.cellsToCarryOver.addAll(OUTFLOWS_RECURRING_BILLS_VALUE_LOCATIONS());
        this.cellsToCarryOver.addAll(OUTFLOWS_RESTRICTED_SAVINGS_LINE_ITEM_LOCATIONS());
        this.cellsToCarryOver.addAll(OUTFLOWS_RESTRICTED_SAVINGS_VALUE_LOCATIONS());
        this.cellsToCarryOver.addAll(OUTFLOWS_OTHER_SAVINGS_LINE_ITEM_LOCATIONS());
        this.cellsToCarryOver.addAll(OUTFLOWS_OTHER_SAVINGS_VALUE_LOCATIONS());
        this.cellsToCarryOver.addAll(OUTFLOWS_SPENDING_MONEY_LINE_ITEM_LOCATIONS());
        this.cellsToCarryOver.addAll(OUTFLOWS_SPENDING_MONEY_VALUE_LOCATIONS());
    }

    @Override
    protected void setLocationsOfFormulaCellsToReevaluate() {
        this.formulaCellsToReevaluate.addAll(PAYCHECK_INCOME_FORMULA_LOCATIONS());
        this.formulaCellsToReevaluate.addAll(ADDITIONAL_RECURRING_INCOME_FORMULA_LOCATIONS());
        this.formulaCellsToReevaluate.addAll(ONE_TIME_INCOME_SUBTOTAL_ROW_NUM());
        this.formulaCellsToReevaluate.addAll(AVAILABLE_INCOME_FORMULA_LOCATIONS());
        this.formulaCellsToReevaluate.addAll(RECURRING_BILLS_FORMULA_LOCATIONS());
        this.formulaCellsToReevaluate.addAll(RESTRICTED_SAVINGS_FORMULA_LOCATIONS());
        this.formulaCellsToReevaluate.addAll(OTHER_SAVINGS_FORMULA_LOCATIONS());
        this.formulaCellsToReevaluate.addAll(SPENDING_MONEY_SUBTOTAL_FORMULA_LOCATIONS());
        this.formulaCellsToReevaluate.addAll(TOTAL_SPENDING_FORMULA_LOCATIONS());
        this.formulaCellsToReevaluate.addAll(EXCESS_AVAILABLE_FUNDS_LOCATIONS());
    }

    private static List<ExcelCellLocation> INFLOWS_ADDITIONAL_RECURRING_INCOME_LINE_ITEM_LOCATIONS() {
        int[] rowsOfCellsToEdit = new int[] {14, 15};
        return getLineItemCellLocationsForRows(rowsOfCellsToEdit);
    }

    private static List<ExcelCellLocation> INFLOWS_ADDITIONAL_RECURRING_INCOME_VALUE_LOCATIONS() {
        int[] rowsOfCellsToEdit = new int[] {14, 15};
        return getValueCellLocationsForRows(rowsOfCellsToEdit);
    }

    private static List<ExcelCellLocation> OUTFLOWS_RECURRING_BILLS_LINE_ITEM_LOCATIONS() {
        int[] rowsOfCellsToEdit = IntStream.range(27, 37).toArray();
        return getLineItemCellLocationsForRows(rowsOfCellsToEdit);    }

    private static List<ExcelCellLocation> OUTFLOWS_RECURRING_BILLS_VALUE_LOCATIONS() {
        int[] rowsOfCellsToEdit = IntStream.range(27, 37).toArray();
        return getValueCellLocationsForRows(rowsOfCellsToEdit);
    }

    private static List<ExcelCellLocation> OUTFLOWS_RESTRICTED_SAVINGS_LINE_ITEM_LOCATIONS() {
        int[] rowsOfCellsToEdit = IntStream.range(40, 50).toArray();
        return getLineItemCellLocationsForRows(rowsOfCellsToEdit);    }

    private static List<ExcelCellLocation> OUTFLOWS_RESTRICTED_SAVINGS_VALUE_LOCATIONS() {
        int[] rowsOfCellsToEdit = IntStream.range(40, 50).toArray();
        return getValueCellLocationsForRows(rowsOfCellsToEdit);
    }

    private static List<ExcelCellLocation> OUTFLOWS_OTHER_SAVINGS_LINE_ITEM_LOCATIONS() {
        int[] rowsOfCellsToEdit = IntStream.range(53, 58).toArray();
        return getLineItemCellLocationsForRows(rowsOfCellsToEdit);    }

    private static List<ExcelCellLocation> OUTFLOWS_OTHER_SAVINGS_VALUE_LOCATIONS() {
        int[] rowsOfCellsToEdit = IntStream.range(53, 58).toArray();
        return getValueCellLocationsForRows(rowsOfCellsToEdit);
    }

    private static List<ExcelCellLocation> OUTFLOWS_SPENDING_MONEY_LINE_ITEM_LOCATIONS() {
        int[] rowsOfCellsToEdit = IntStream.range(61, 71).toArray();
        return getLineItemCellLocationsForRows(rowsOfCellsToEdit);    }

    private static List<ExcelCellLocation> OUTFLOWS_SPENDING_MONEY_VALUE_LOCATIONS() {
        int[] rowsOfCellsToEdit = IntStream.range(61, 71).toArray();
        return getValueCellLocationsForRows(rowsOfCellsToEdit);
    }

    private static List<ExcelCellLocation> PAYCHECK_INCOME_FORMULA_LOCATIONS() {
        int[] rowsOfCellsToEdit = IntStream.range(10, 12).toArray();
        ExcelColumn[] columnsOfCellsToEdit = new ExcelColumn[] {
                // RECURRING FORMULA ROWS TAKEN FROM PAYCHECK SHEET
                ExcelColumn.H, ExcelColumn.I, ExcelColumn.J, ExcelColumn.K, ExcelColumn.L, ExcelColumn.M,

                // CALCULATION COLUMN NUMS
                ExcelColumn.N, ExcelColumn.O, ExcelColumn.P,

                // ACTUAL MONTHLY VALUE COLUMN NUMS
                ExcelColumn.Q, ExcelColumn.R, ExcelColumn.S, ExcelColumn.T, ExcelColumn.U, ExcelColumn.V, ExcelColumn.W,
                ExcelColumn.X, ExcelColumn.Y, ExcelColumn.Z, ExcelColumn.AA, ExcelColumn.AB
        };
        return createCellListPermutations(columnsOfCellsToEdit, rowsOfCellsToEdit, ExcelCellLocation.VALUE_TYPE.NUMERIC);
    }

    private static List<ExcelCellLocation> ADDITIONAL_RECURRING_INCOME_FORMULA_LOCATIONS() {
        int[] rowsOfCellsToEdit = IntStream.range(14, 17).toArray();
        return getFormulaCellLocationsForRows(rowsOfCellsToEdit);
    }

    private static List<ExcelCellLocation> ONE_TIME_INCOME_SUBTOTAL_ROW_NUM() {
        int[] rowsOfCellsToEdit = IntStream.range(19, 22).toArray();
        return getFormulaCellLocationsForRows(rowsOfCellsToEdit);
    }

    private static List<ExcelCellLocation> AVAILABLE_INCOME_FORMULA_LOCATIONS() {
        int[] rowsOfCellsToEdit = new int[]{23};
        return getFormulaCellLocationsForRows(rowsOfCellsToEdit);
    }

    private static List<ExcelCellLocation> RECURRING_BILLS_FORMULA_LOCATIONS() {
        int[] rowsOfCellsToEdit = IntStream.range(27, 38).toArray();
        return getFormulaCellLocationsForRows(rowsOfCellsToEdit);
    }

    private static List<ExcelCellLocation> RESTRICTED_SAVINGS_FORMULA_LOCATIONS() {
        int[] rowsOfCellsToEdit = IntStream.range(40, 51).toArray();
        return getFormulaCellLocationsForRows(rowsOfCellsToEdit);
    }

    private static List<ExcelCellLocation> OTHER_SAVINGS_FORMULA_LOCATIONS() {
        int[] rowsOfCellsToEdit = IntStream.range(53, 59).toArray();
        return getFormulaCellLocationsForRows(rowsOfCellsToEdit);
    }

    private static List<ExcelCellLocation> SPENDING_MONEY_SUBTOTAL_FORMULA_LOCATIONS() {
        int[] rowsOfCellsToEdit = IntStream.range(61, 72).toArray();
        return getFormulaCellLocationsForRows(rowsOfCellsToEdit);
    }

    private static List<ExcelCellLocation> TOTAL_SPENDING_FORMULA_LOCATIONS() {
        int[] rowsOfCellsToEdit = new int[]{73};
        return getFormulaCellLocationsForRows(rowsOfCellsToEdit);
    }

    private static List<ExcelCellLocation> EXCESS_AVAILABLE_FUNDS_LOCATIONS() {
        int[] rowsOfCellsToEdit = new int[]{75};
        return getFormulaCellLocationsForRows(rowsOfCellsToEdit);
    }

    private static List<ExcelCellLocation> getLineItemCellLocationsForRows(int[] rowsOfCellsToEdit) {
        ExcelColumn[] columnsOfCellsToEdit = new ExcelColumn[] {ExcelColumn.C};
        return createCellListPermutations(columnsOfCellsToEdit, rowsOfCellsToEdit, ExcelCellLocation.VALUE_TYPE.STRING);
    }

    private static List<ExcelCellLocation> getValueCellLocationsForRows(int[] rowsOfCellsToEdit) {
        ExcelColumn[] columnsOfCellsToEdit = new ExcelColumn[] {
                ExcelColumn.H, ExcelColumn.I, ExcelColumn.J, ExcelColumn.K, ExcelColumn.L, ExcelColumn.M
        };
        return createCellListPermutations(columnsOfCellsToEdit, rowsOfCellsToEdit, ExcelCellLocation.VALUE_TYPE.NUMERIC);
    }

    private static List<ExcelCellLocation> getFormulaCellLocationsForRows(int[] rowsOfCellsToEdit) {
        ExcelColumn[] columnsOfCellsToEdit = new ExcelColumn[] {
                // CALCULATION COLUMN NUMS
                ExcelColumn.N, ExcelColumn.O, ExcelColumn.P,

                // ACTUAL MONTHLY VALUE COLUMN NUMS
                ExcelColumn.Q, ExcelColumn.R, ExcelColumn.S, ExcelColumn.T, ExcelColumn.U, ExcelColumn.V, ExcelColumn.W,
                ExcelColumn.X, ExcelColumn.Y, ExcelColumn.Z, ExcelColumn.AA, ExcelColumn.AB
        };
        return createCellListPermutations(columnsOfCellsToEdit, rowsOfCellsToEdit, ExcelCellLocation.VALUE_TYPE.NUMERIC);
    }
}
