package cn.keats.excel.entity;

import lombok.Data;

import java.util.Date;

/**
 * 功能：基础实体
 *
 * @author kangshuai@gridsum.com
 * @date 2021/6/16 14:22
 */
@Data
public class BaseEntity {

    private Integer deleteFlag;

    private String createUser;

    private Date createTime;

    private String updateUser;

    private Date updateTime;
}
