package cn.keats.excel.controller;

import cn.afterturn.easypoi.excel.ExcelExportUtil;
import cn.afterturn.easypoi.excel.entity.ExportParams;
import cn.afterturn.easypoi.excel.entity.enmus.ExcelType;
import cn.keats.excel.entity.Human;
import cn.keats.excel.helper.ExcelStyleHelper;
import cn.keats.excel.helper.FieldOrderMappingHelper;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;

/**
 * 功能：导出 Excel
 *
 * @author kangshuai@gridsum.com
 * @date 2021/6/16 14:29
 */
@RestController
public class ExcelExportController {

    @GetMapping("excel")
    public void excelExport(HttpServletResponse response) throws Exception {
        List<Human> humanList = new ArrayList<>();

        doWriteListToResponse(humanList, Human.class, response, "测试 Sheet", "测试 Excel.xls");
    }


    /**
     * 功能：将结果写入输出流
     *
     * @author kangshuai@gridsum.com
     * @date 2021/4/14 14:46
     */
    public <T> void doWriteListToResponse(List<T> list, Class<T> exportType, HttpServletResponse response, String sheetName, String excelName) throws IOException {
        ExportParams ex = new ExportParams(null, sheetName, ExcelType.HSSF);
        // 创建导出对象
        Workbook workbook = ExcelExportUtil.exportExcel(ex, exportType, list);
        // 初始化工具类
        HashMap<String, Integer> map = new FieldOrderMappingHelper<>(exportType).getFieldAndOrderMap();
        ExcelStyleHelper styleHelper = new ExcelStyleHelper(workbook);
        // 添加规则
        styleHelper.setValidation(exportType, map);
        // 写入输出流，忽略此处硬编码
        response.setHeader("Content-Disposition", "attachment;filename=" + new String(excelName.getBytes(StandardCharsets.UTF_8), StandardCharsets.ISO_8859_1));
        response.setCharacterEncoding("UTF-8");
        response.setContentType("application/x-download");
        workbook.write(response.getOutputStream());
    }
}
