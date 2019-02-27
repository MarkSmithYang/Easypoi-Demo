package com.yb.easypoi.model;

import cn.afterturn.easypoi.excel.annotation.Excel;
import cn.afterturn.easypoi.excel.annotation.ExcelCollection;
import cn.afterturn.easypoi.excel.annotation.ExcelEntity;
import cn.afterturn.easypoi.excel.annotation.ExcelTarget;

import javax.persistence.Entity;
import javax.persistence.Id;
import java.io.Serializable;
import java.util.List;
import java.util.UUID;

/**
 * @author yangbiao
 * @Description:课程类
 * @date 2018/10/29
 */
@ExcelTarget("course")
public class Course implements Serializable {
    private static final long serialVersionUID = 989165736995970616L;

    /**
     * id
     */
    private String id;

    /**
     * 课程名称
     */
    @Excel(name = "课程名称", orderNum = "1", needMerge = true,mergeVertical = true)
    private String courseName;

    /**
     * 教师信息
     */
    @ExcelEntity(id="abb")
    private Teacher teacher;

    /**
     * 学生集合
     */
    //这里的orderNum实测仅仅只是排列学生集合的顺序而已,
    //一般正常按正常的顺序设置即可,也可以用比较大的数字,
    //例如我的是100,实测2,3,4和一般都一样的,但是不设置顺序就会乱
    @ExcelCollection(name = "学生信息", orderNum = "100")
    private List<Student> list;

    public Course() {
        this.id = UUID.randomUUID().toString().replaceAll("-", "");
    }

    public Course(String courseName) {
        this.courseName = courseName;
        this.id = UUID.randomUUID().toString().replaceAll("-", "");
    }

    public String getId() {
        return id;
    }

    public void setId(String id) {
        this.id = id;
    }

    public String getCourseName() {
        return courseName;
    }

    public void setCourseName(String courseName) {
        this.courseName = courseName;
    }

    public Teacher getTeacher() {
        return teacher;
    }

    public void setTeacher(Teacher teacher) {
        this.teacher = teacher;
    }

    public List<Student> getList() {
        return list;
    }

    public void setList(List<Student> list) {
        this.list = list;
    }

    @Override
    public String toString() {
        return "Course{" +
                "id='" + id + '\'' +
                ", courseName='" + courseName + '\'' +
                ", teacher=" + teacher +
                ", list=" + list +
                '}';
    }
}