package cn.keats.excel.helper;


import cn.afterturn.easypoi.excel.annotation.Excel;
import cn.keats.excel.entity.FiledAndOrder;

import java.lang.reflect.Field;
import java.util.*;

/**
 * 功能：获取某个 pojo 类的字段和顺序、以及字段和 excel 列名的对应关系
 *
 * @author kangshuai@gridsum.com
 * @date 2021/4/9 11:50
 */
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
