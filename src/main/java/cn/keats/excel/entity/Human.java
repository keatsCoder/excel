package cn.keats.excel.entity;

import cn.afterturn.easypoi.excel.annotation.Excel;
import lombok.Data;

/**
 * 功能：人类-实体
 *
 * @author kangshuai@gridsum.com
 * @date 2021/6/16 14:22
 */
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
