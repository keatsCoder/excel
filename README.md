# 一句话总结

Excel 导出、导入时，根据注解自动添加单元格验证规则，避免用户因填写错误的枚举字段而反复修改 Excel

# 需求背景

对于 Java Web 项目，总是不可避免的出现 Excel 导入、导出的需求，而 Excel 导入导出时，枚举字段和枚举值的映射是非常常见的一种情况

例如：下面这张示例中的性别列

数据库表结构：

![image-20210616140903264](https://img2020.cnblogs.com/blog/1654189/202106/1654189-20210616154617365-826467944.png)

Excel 中用户需要输入：男，女，未知

![image-20210616140615962](https://img2020.cnblogs.com/blog/1654189/202106/1654189-20210616154617164-1013098015.png)

常见的 Excel 框架都已经覆盖了枚举映射的功能，例如：[EasyPOI](http://doc.wupaas.com/docs/easypoi/easypoi-1c0u4mo8p4ro8) 

但是这种操作方式对于用户来说，并不是很方便，试想一下：假如用户在性别列输入了：男性，最终的结果一般就是程序抛出异常，用户得到提示：性别输入有误，贴心的开发者可能会加上：请输入 男 女 未知，做的更好一些的 可能在列头添加标签提示：该列仅能输入 男 女 未知，但是这种弱限制也无法从根本上解决问题

![image-20210616141205671](https://img2020.cnblogs.com/blog/1654189/202106/1654189-20210616154616945-856745608.png)

更好一点的解决方案是：利用 Excel 的数据验证功能，把单元格加上规则校验，让用户只能输入正确的枚举值，避免因一次输入错误而反复返工，浪费用户的时间和好心情

![image-20210616141620860](https://img2020.cnblogs.com/blog/1654189/202106/1654189-20210616154616655-1302228324.png)

当用户输入了非枚举值之后，Excel 会提示用户输入不合规，禁止用户保存

![image-20210616141707359](https://img2020.cnblogs.com/blog/1654189/202106/1654189-20210616154616263-1820333829.png)

这样的交互就能从源头保证用户输入正确的值

那这么友好的设计，在 Java 中如何能方便且可扩展性更强的实现呢？

# 需求实现

我这边的实现是基于 [EasyPOI](http://doc.wupaas.com/docs/easypoi/easypoi-1c0u4mo8p4ro8) + 注解(EasyPOI 转换映射关系注解复用) + 反射 实现的，解决了以上需求痛点的同时，可以满足代码一处修改，多个功能都生效的目的

## 代码仓库

[GayHub](https://github.com/keatsCoder/excel)

## 实体类

@Excel 注解中的 replace 属性，该属性是 EasyPOI 用来做字段映射的，我这里复用他做 Excel 验证的可选项，另外一个就是 orderNum 属性，用该值来自动获取某个字段在 Excel 中的列的位置

```java
@Data
public class Human extends BaseEntity {

    private Long id;

    @Excel(name = "姓名", orderNum = "1", width = 15)
    private String name;

    @Excel(name = "年龄", orderNum = "2", width = 15)
    private Integer age;

    @Excel(name = "性别", replace = {"男_1", "女_2", "未知_3"}, orderNum = "3", width = 15)
    private Integer gender;
}
```

## 获取列名和列位置的映射

该类在初始化时，需要指定当前导出 Excel 对应的实体类的类类型，然后通过遍历类中字段的注解，生成字段和列排序(位置)的映射关系

```java
public class FieldOrderMappingHelper<T> {
    /**
     * 支持的最大字段数
     */
    private final static int MAX_LIST_SIZE = 26;

    public FieldOrderMappingHelper(Class<T> pojo) {
        this.pojo = pojo;
        initMap();
    }

    /**
     * 解析注释的 pojo 对象
     */
    private Class<T> pojo;

    /**
     * 字段和序号的映射关系
     */
    private HashMap<String, Integer> fieldAndOrderMap;


    /**
     * 功能：初始化类的字段内容，建立字段和序号以及字段和 excel 列名的映射关系
     *
     * @author kangshuai@gridsum.com
     * @date 2021/4/9 12:06
     */
    private void initMap() {
        HashMap<String, Integer> fieldAndOrderMap = new HashMap<>(16);
        HashSet<Integer> existOrderNumSet = new HashSet<>(16);

        List<FiledAndOrder> list = new ArrayList<>();
        list = initList(list, pojo);
        if (list.size() > MAX_LIST_SIZE) {
            throw new RuntimeException(pojo.getName() + "目前最大支持 26 个字段，26+ 需要改代码");
        }

        // 排序
        list.sort(Comparator.comparing(FiledAndOrder::getOrder));

        for (int i = 0; i < list.size(); i++) {
            if (existOrderNumSet.contains(list.get(i).getOrder())) {
                throw new RuntimeException(pojo.getName() + "类内部或与父类字段中存在重复的 excel 排序，请修改");
            }
            existOrderNumSet.add(list.get(i).getOrder());
            fieldAndOrderMap.put(list.get(i).getFiledName(), i);
        }
        this.fieldAndOrderMap = fieldAndOrderMap;
    }

    /**
     * 功能：初始化类的字段信息，转换成 ArrayList
     *
     * @return java.util.List<com.gridsum.ad.ooh.project.entity.FiledAndOrder>
     * @author kangshuai@gridsum.com
     * @date 2021/4/9 12:09
     */
    private List<FiledAndOrder> initList(List<FiledAndOrder> list, Class<?> pojoClass) {
        if (Object.class.equals(pojoClass)) {
            return list;
        }
        Field[] fields = pojoClass.getDeclaredFields();
        for (Field f : fields) {
            // 找到所有加了 Excel 注解的字段
            Excel annotation = f.getAnnotation(Excel.class);
            if (annotation == null) {
                continue;
            }
            // 过滤隐藏行
            if (annotation.isColumnHidden()) {
                continue;
            }
            FiledAndOrder filedAndOrder = new FiledAndOrder(f.getName(), Integer.parseInt(annotation.orderNum()));
            list.add(filedAndOrder);
        }
        // 递归查找父类
        Class<?> superclass = pojoClass.getSuperclass();
        return initList(list, superclass);
    }

    public HashMap<String, Integer> getFieldAndOrderMap() {
        return fieldAndOrderMap;
    }
}
```

## 设置验证规则

setValidation 方法有两个参数，第一个是导出 Excel 对应的实体类的类类型，第二个是 FieldOrderMappingHelper.getFieldAndOrderMap() 获取到的列名和排序映射，该类通过反射字段上的注解，自动为生成的 workbook 添加验证规则

```java
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
```

## 示例导出代码

控制层代码如下

```java
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
```

