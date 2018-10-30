package com.yb.easypoi.model;

import cn.afterturn.easypoi.excel.annotation.Excel;
import cn.afterturn.easypoi.excel.annotation.ExcelEntity;

import java.io.Serializable;
import java.time.LocalDate;
import java.util.UUID;

/**
 * @author yangbiao
 * @Description:学生实体类--这个需要配合实体类使用的方式最好不要用jdbc来获取数据,使用jpa为最好
 * @date 2018/10/23
 */
public class Student implements Serializable {
    private static final long serialVersionUID = -3944779670897279495L;

    /**
     * 入学编号
     */
    @Excel(name = "入学编号", orderNum = "1", needMerge = true)
    private String id;

    /**
     * 姓名
     */
    @Excel(name = "姓名", orderNum = "2", needMerge = true)
    private String name;

    /**
     * 年龄
     */
    @Excel(name = "年龄", orderNum = "3", needMerge = true,mergeVertical = true)
    private Integer age;

    /**
     * 班级
     */
    @Excel(name = "班级", orderNum = "4", mergeVertical = true)
    private String className;

    /**
     * 入学时间
     */
    @Excel(name = "入学时间", orderNum = "5", needMerge = true, mergeVertical = true)
    private LocalDate joinTime;

    public Student() {
        //构造时初始化id和入学时间
        this.id = UUID.randomUUID().toString().replaceAll("-:", "");
        this.joinTime = LocalDate.now();
    }

    public Student(String name, Integer age, String className) {
        this.name = name;
        this.age = age;
        this.className = className;
        //构造时初始化id和入学时间
        this.id = UUID.randomUUID().toString().replaceAll("-:", "");
        this.joinTime = LocalDate.now();
    }

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

    public Integer getAge() {
        return age;
    }

    public void setAge(Integer age) {
        this.age = age;
    }

    public String getClassName() {
        return className;
    }

    public void setClassName(String className) {
        this.className = className;
    }

    public LocalDate getJoinTime() {
        return joinTime;
    }

    public void setJoinTime(LocalDate joinTime) {
        this.joinTime = joinTime;
    }

    @Override
    public String toString() {
        return "Student{" +
                "id='" + id + '\'' +
                ", name='" + name + '\'' +
                ", age=" + age +
                ", className='" + className + '\'' +
                ", joinTime=" + joinTime +
                '}';
    }
}
