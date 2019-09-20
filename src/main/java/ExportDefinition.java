public class ExportDefinition {
    private String title; // 标题
    private String field; // 字段
    private int rowIndex; // 所在的行
    private int cellIndex; // 所在的列
    private String mainDict; // 主字典-用于加载主字典的数据
    private String subDict; // 子字典-用于加载subField的数据
    private String subField; // 即需要级联的字典
    private String refName; // 主字段所在的位置
    private String point; // 标题的坐标
    private boolean validate;// 是否设置数据的有限性


    public String getTitle() {
        return title;
    }

    public void setTitle(String title) {
        this.title = title;
    }

    public String getField() {
        return field;
    }

    public void setField(String field) {
        this.field = field;
    }

    public int getRowIndex() {
        return rowIndex;
    }

    public void setRowIndex(int rowIndex) {
        this.rowIndex = rowIndex;
    }

    public int getCellIndex() {
        return cellIndex;
    }

    public void setCellIndex(int cellIndex) {
        this.cellIndex = cellIndex;
    }

    public String getMainDict() {
        return mainDict;
    }

    public void setMainDict(String mainDict) {
        this.mainDict = mainDict;
    }

    public String getSubDict() {
        return subDict;
    }

    public void setSubDict(String subDict) {
        this.subDict = subDict;
    }

    public String getSubField() {
        return subField;
    }

    public void setSubField(String subField) {
        this.subField = subField;
    }

    public String getRefName() {
        return refName;
    }

    public void setRefName(String refName) {
        this.refName = refName;
    }

    public String getPoint() {
        return point;
    }

    public void setPoint(String point) {
        this.point = point;
    }

    public boolean isValidate() {
        return validate;
    }

    public void setValidate(boolean validate) {
        this.validate = validate;
    }
    public ExportDefinition(String title, String field, String mainDict, String subDict, String subField) {
        this.title = title;
        this.field = field;
        this.mainDict = mainDict;
        this.subDict = subDict;
        this.subField = subField;
    }
}
