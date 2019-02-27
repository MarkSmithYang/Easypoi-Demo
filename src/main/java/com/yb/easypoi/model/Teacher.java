package com.yb.easypoi.model;

import cn.afterturn.easypoi.excel.annotation.Excel;
import cn.afterturn.easypoi.excel.annotation.ExcelTarget;

import javax.persistence.Entity;
import javax.persistence.Id;
import java.io.Serializable;
import java.util.UUID;

/**
 * @author yangbiao
 * @Description:教师类
 * @date 2018/10/29
 */
@Entity
@ExcelTarget("teacher")
public class Teacher implements Serializable {
    private static final long serialVersionUID = 1095168552207599306L;

    /**
     * id
     */
    @Id
    private String id;

    /**
     * 教师姓名
     */
    @Excel(name = "教师姓名", orderNum = "1", needMerge = true,mergeVertical = true)
    private String name;

    public String getId() {
        return id;
    }

    public void setId(String id) {
        this.id = id;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public Teacher() {
        this.id = UUID.randomUUID().toString().replaceAll("-", "");
    }

    public Teacher(String name) {
        this.name = name;
        this.id = UUID.randomUUID().toString().replaceAll("-", "");
    }

    @Override
    public String toString() {
        return "Teacher{" +
                "id='" + id + '\'' +
                ", name='" + name + '\'' +
                '}';
    }
}
