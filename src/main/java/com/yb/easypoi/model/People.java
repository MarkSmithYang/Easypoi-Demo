package com.yb.easypoi.model;

import cn.afterturn.easypoi.excel.annotation.Excel;
import cn.afterturn.easypoi.excel.entity.result.ExcelVerifyHandlerResult;
import cn.afterturn.easypoi.handler.inter.IExcelDataModel;
import cn.afterturn.easypoi.handler.inter.IExcelModel;
import cn.afterturn.easypoi.handler.inter.IExcelVerifyHandler;
import org.apache.poi.ss.formula.functions.T;

import javax.validation.constraints.NotBlank;
import java.io.Serializable;

/**
 * @author yangbiao
 * @Description:
 * @date 2018/10/30
 */
public class People implements Serializable, IExcelModel, IExcelDataModel, IExcelVerifyHandler<T> {
    private static final long serialVersionUID = 5573829739286021262L;

    /**
     * id
     */
    @Excel(name = "编号")
    @NotBlank(message = "编号不能为空")
    private String id;

    /**
     * 姓名
     */
//    @Pattern(regexp = "^[\\u4E00-\\u9FA5]{3}$",message = "姓名不是中文")
    @Excel(name = "姓名", needMerge = true, mergeVertical = true)
//    @Pattern(regexp = "^[j]{1}[a]{1}[c]{1}[k]{1}$",message = "姓名不符合要求")
    private String name;

    /**
     * 住址
     */
//    @Pattern(regexp = "^[\\u4E00-\\u9FA5]{3}$",message = "住址最低需要三个中文字")
    @Excel(name = "住址", needMerge = true, mergeVertical = true)
    private String address;

    public People() {
    }

    public People(String name, String address) {
        this.name = name;
        this.address = address;
    }

    @Override
    public String toString() {
        return "People{" +
                "id='" + id + '\'' +
                ", name='" + name + '\'' +
                ", address='" + address + '\'' +
                '}';
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

    public String getAddress() {
        return address;
    }

    public void setAddress(String address) {
        this.address = address;
    }

    //000000000000000000000000000000000000000000000000000000000000000000000000000000000000
    /**
     * 错误信息
     */
    private String errorMsg;

    /**
     * 所在行(行号)
     */
    private int rowNum;


    @Override
    public int getRowNum() {
        return rowNum;
    }

    @Override
    public void setRowNum(int rowNum) {
        this.rowNum = rowNum;
        verifyHandler(null);
    }

    @Override
    public String getErrorMsg() {
        return errorMsg;
    }

    @Override
    public void setErrorMsg(String s) {
        this.errorMsg = errorMsg;
    }

    @Override
    public ExcelVerifyHandlerResult verifyHandler(T t) {
        if ("jack".equals(this.name)) {
            System.err.println("ok");
        }
        return null;
    }
}