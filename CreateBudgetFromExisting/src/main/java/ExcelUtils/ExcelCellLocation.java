package ExcelUtils;

public class ExcelCellLocation {
    private ExcelColumn column;
    private int row;
    private VALUE_TYPE valueType;
    public enum VALUE_TYPE {
        STRING,
        NUMERIC
    }

    public ExcelCellLocation(ExcelColumn col, int row, VALUE_TYPE valueType) {
        this.column = col;
        this.row = row - 1;
        this.valueType = valueType;
    }

    public int getColumn() {
        return this.column.getValue();
    }

    public int getRow() {
        return this.row;
    }

    public VALUE_TYPE getValueType() {
        return this.valueType;
    }
}