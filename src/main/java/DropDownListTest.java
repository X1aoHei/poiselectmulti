import org.apache.poi.hssf.usermodel.DVConstraint;
import org.apache.poi.hssf.usermodel.HSSFDataValidation;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.ss.util.CellReference;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Objects;

public class DropDownListTest {
    private static final String DICT_SHEET = "DICT_SHEET";

    public static void main(String[] args) throws IOException {
        // 1.准备需要生成excel模板的数据
        List<ExportDefinition> edList = new ArrayList<>(2);
        //edList.add(new ExportDefinition("生活用品", "xx", null, null, null));
        edList.add(new ExportDefinition("任务来源大类", "dl", "car-dict", "fruit-dict", "nr"));
        edList.add(new ExportDefinition("任务来源内容", "nr", "fruit-dict", "", ""));
        //edList.add(new ExportDefinition("测试", "yy", "t-dict", null, null));

        // 2.生成导出模板
        Workbook wb = new HSSFWorkbook();
        Sheet sheet = createExportSheet(edList, wb);

        // 3.创建数据字典sheet页
        createDictSheet(edList, wb);

        // 4.设置数据有效性
        setDataValidation(edList, sheet);

        // 5.保存excel到本地
        OutputStream os = new FileOutputStream("d:/4.xls");
        wb.write(os);

        System.out.println("模板生成成功.");
    }
    public static void createDataValidateSubList(Sheet sheet, ExportDefinition ed) {
        int rowIndex = ed.getRowIndex();
        CellRangeAddressList cal;
        DVConstraint constraint;
        CellReference cr;
        DataValidation dataValidation;
        System.out.println(ed);
        for (int i = 0; i < 100; i++) {
            int tempRowIndex = ++rowIndex;
            cal = new CellRangeAddressList(tempRowIndex, tempRowIndex, ed.getCellIndex(), ed.getCellIndex());
            cr = new CellReference(rowIndex, ed.getCellIndex() - 1, true, true);
            constraint = DVConstraint.createFormulaListConstraint("INDIRECT(" + cr.formatAsString() + ")");
            dataValidation = new HSSFDataValidation(cal, constraint);
            dataValidation.setSuppressDropDownArrow(false);
            dataValidation.createPromptBox("操作提示", "请选择下拉选中的值");
            dataValidation.createErrorBox("错误提示", "请从下拉选中选择，不要随便输入");
            sheet.addValidationData(dataValidation);
        }
    }

    /**
     * @param edList
     * @param sheet
     */
    private static void setDataValidation(List<ExportDefinition> edList, Sheet sheet) {
        for (ExportDefinition ed : edList) {
            if (ed.isValidate()) {// 说明是下拉选
                DVConstraint constraint = DVConstraint.createFormulaListConstraint(ed.getField());
                if (null == ed.getRefName()) {// 说明是一级下拉选
                    createDataValidate(sheet, ed, constraint);
                } else {// 说明是二级下拉选
                    createDataValidateSubList(sheet, ed);
                }
            }
        }
    }

    /**
     * @param sheet
     * @param ed
     * @param constraint
     */
    private static void createDataValidate(Sheet sheet, ExportDefinition ed, DVConstraint constraint) {
        CellRangeAddressList regions = new CellRangeAddressList(ed.getRowIndex() + 1, ed.getRowIndex() + 100, ed.getCellIndex(), ed.getCellIndex());
        DataValidation dataValidation = new HSSFDataValidation(regions, constraint);
        dataValidation.setSuppressDropDownArrow(false);
        // 设置提示信息
        dataValidation.createPromptBox("操作提示", "请选择下拉选中的值");
        // 设置输入错误信息
        dataValidation.createErrorBox("错误提示", "请从下拉选中选择，不要随便输入");
        sheet.addValidationData(dataValidation);
    }

    /**
     * @param edList
     * @param wb
     */
    private static void createDictSheet(List<ExportDefinition> edList, Workbook wb) {
        Sheet sheet = wb.createSheet(DICT_SHEET);
        RowCellIndex rci = new RowCellIndex(0, 0);
        for (ExportDefinition ed : edList) {
            String mainDict = ed.getMainDict();
            if (null != mainDict && null == ed.getRefName()) {// 是第一个下拉选
                List<String> mainDictList = (List<String>) DictData.getDict(mainDict);
                String refersToFormula = createDictAndReturnRefFormula(sheet, rci, mainDictList);
                // 创建 命名管理
                createName(wb, ed.getField(), refersToFormula);
                ed.setValidate(true);
            }
            if (null != mainDict && null != ed.getSubDict() && null != ed.getSubField()) {// 联动时加载ed.getSubField()的数据
                ExportDefinition subEd = fiterByField(edList, ed.getSubField());// 获取需要级联的那个字段
                if (null == subEd) {
                    continue;
                }
                subEd.setRefName(ed.getPoint());// 保存主下拉选的位置
                subEd.setValidate(true);
                Map<String, List<String>> subDictListMap = (Map<String, List<String>>) DictData.getDict(ed.getSubDict());
                for (Map.Entry<String, List<String>> entry : subDictListMap.entrySet()) {
                    String refersToFormula = createDictAndReturnRefFormula(sheet, rci, entry.getValue());
                    // 创建 命名管理
                    createName(wb, entry.getKey(), refersToFormula);
                }
            }
        }
    }

    /**
     * @param sheet
     * @param rci
     * @param
     * @return
     */
    private static String createDictAndReturnRefFormula(Sheet sheet, RowCellIndex rci, List<String> datas) {
        Row row = sheet.createRow(rci.incrementRowIndexAndGet());
        rci.setCellIndex(0);
        int startRow = rci.getRowIndex();
        int startCell = rci.getCellIndex();
        for (String dict : datas) {
            row.createCell(rci.incrementCellIndexAndGet()).setCellValue(dict);
        }
        int endRow = rci.getRowIndex();
        int endCell = rci.getCellIndex();
        String startName = new CellReference(DICT_SHEET, startRow, startCell, true, true).formatAsString();
        String endName = new CellReference(endRow, endCell, true, true).formatAsString();
        String refersToFormula = startName + ":" + endName;
        System.out.println(refersToFormula);
        return refersToFormula;
    }

    /**
     * @param wb
     * @param nameName
     *            表示命名管理的名字
     * @param refersToFormula
     */
    private static void createName(Workbook wb, String nameName, String refersToFormula) {
        Name name = wb.createName();
        name.setNameName(nameName);
        name.setRefersToFormula(refersToFormula);
    }

    private static ExportDefinition fiterByField(List<ExportDefinition> edList, String field) {
        for (ExportDefinition ed : edList) {
            if (Objects.equals(ed.getField(), field)) {
                return ed;
            }
        }
        return null;
    }

    /**
     * @param edList
     * @param wb
     */
    private static Sheet createExportSheet(List<ExportDefinition> edList, Workbook wb) {
        Sheet sheet = wb.createSheet("导出模板");
        RowCellIndex rci = new RowCellIndex(0, 0);
        Row row = sheet.createRow(rci.getRowIndex());
        CellReference cr = null;
        for (ExportDefinition ed : edList) {
            row.createCell(rci.incrementCellIndexAndGet()).setCellValue(ed.getTitle());
            ed.setRowIndex(rci.getRowIndex());
            ed.setCellIndex(rci.getCellIndex());
            cr = new CellReference(ed.getRowIndex() + 1, ed.getCellIndex(), true, true);
            ed.setPoint(cr.formatAsString());
        }
        return sheet;
    }
}
