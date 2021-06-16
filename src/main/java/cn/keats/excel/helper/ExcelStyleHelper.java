package cn.keats.excel.helper;

import cn.afterturn.easypoi.excel.annotation.Excel;
import org.apache.poi.hssf.usermodel.DVConstraint;
import org.apache.poi.hssf.usermodel.HSSFDataValidation;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddressList;

import java.lang.reflect.Field;
import java.util.HashMap;

/**
 * 功能：修改 excel 模板样式
 *
 * @author kangshuai@gridsum.com
 * @date 2021/4/9 15:01
 */
public class ExcelStyleHelper {

    /**
     * 添加列值验证的最小行
     */
    public static final int EXCEL_VALID_ROW_MIN = 1;
    /**
     * 添加列值验证的最大行
     */
    public static final int EXCEL_VALID_ROW_MAX = (2 << 15) - 1;
    /**
     * Excel 对象
     */
    private Workbook workbook;
    /**
     * Sheet 页，默认取第一个 sheet 页
     */
    private Sheet sheet;

    public ExcelStyleHelper(Workbook workbook) {
        this.workbook = workbook;
        this.sheet = workbook.getSheetAt(0);
    }
    
    /**
     * 功能：单元格添加下拉框，仅支持 xls
     *
     * @author kangshuai@gridsum.com
     * @date 2021/4/8 18:55
     */
    public void setValidation(Class<?> pojoClass, HashMap<String, Integer> map) {
        // 递归到 Object 就停下
        if (Object.class.equals(pojoClass)) {
            return;
        }
        // 获取所有的字段
        Field[] fields = pojoClass.getDeclaredFields();
        for (Field field : fields) {
            Excel annotation = field.getAnnotation(Excel.class);
            if (annotation == null) {
                continue;
            }
            String[] replace = annotation.replace();
            if (replace.length == 0) {
                continue;
            }
            String[] textList = new String[replace.length];
            for (int i = 0; i < replace.length; i++) {
                textList[i] = replace[i].split("_")[0];
            }
            // 根据字段名获取他在 excel 中的列数（结合 excel 注解中的排序）
            Integer col = map.get(field.getName());
            setValid(textList, col, col);
        }
        // 递归父类的注解
        Class<?> superclass = pojoClass.getSuperclass();
        setValidation(superclass, map);
    }

    /**
     * 功能：设置验证区间
     *
     * @author kangshuai@gridsum.com
     * @date 2021/4/9 15:11
     */
    private void setValid(String[] textList, int firstCol, int endCol) {
        // 设置数据有效性加载在哪个单元格上,四个参数分别是：起始行、终止行、起始列、终止列
        CellRangeAddressList regions = new CellRangeAddressList(EXCEL_VALID_ROW_MIN, EXCEL_VALID_ROW_MAX, firstCol, endCol);
        // 加载下拉列表内容
        DVConstraint constraint = DVConstraint.createExplicitListConstraint(textList);
        // 数据有效性对象
        HSSFDataValidation dataList = new HSSFDataValidation(regions, constraint);
        sheet.addValidationData(dataList);
    }
}
