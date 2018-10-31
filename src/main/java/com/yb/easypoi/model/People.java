package com.yb.easypoi.model;

import cn.afterturn.easypoi.excel.annotation.Excel;
import cn.afterturn.easypoi.excel.entity.result.ExcelVerifyHandlerResult;
import cn.afterturn.easypoi.handler.inter.IExcelDataModel;
import cn.afterturn.easypoi.handler.inter.IExcelModel;
import cn.afterturn.easypoi.handler.inter.IExcelVerifyHandler;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.formula.functions.T;

import javax.validation.constraints.NotBlank;
import javax.validation.constraints.Pattern;
import java.io.Serializable;

/**
 * @author yangbiao
 * @Description:
 * @date 2018/10/30
 */
public class People implements Serializable, IExcelModel, IExcelDataModel, IExcelVerifyHandler<People> {
    private static final long serialVersionUID = 5573829739286021262L;

    /**
     * id
     */
    @Excel(name = "编号", isImportField = "true")
    @NotBlank(message = "编号不能为空")
    private String id;

    /**
     * 姓名
     */
    @Pattern(regexp = "^[\\u4E00-\\u9FA5]*$", message = "姓名不是中文")
    @Excel(name = "姓名", needMerge = true, mergeVertical = true)
    private String name;

    /**
     * 住址
     */
    @Pattern(regexp = "^[\\u4E00-\\u9FA5]{3}$", message = "住址最低需要三个中文字")
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
    @Excel(name = "错误信息", needMerge = true, mergeVertical = true)
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
        //因为每条数据都要走这个方法,在这里调用需要校验的方法是很好的.
        verifyHandler(this);
    }

    @Override
    public String getErrorMsg() {
        return errorMsg;
    }

    @Override
    public void setErrorMsg(String errorMsg) {
        //把实体属性全都做非空判断,只有需要的数据全部为空的时候才能判定此行为无效信息行(就是不含有我们需要信息的行)
        if (StringUtils.isBlank(id) && StringUtils.isBlank(name) && StringUtils.isBlank(address)) {
            //清空导入api对无效单元格行的校验信息-->为了后面通过这个的非空判断获取有效单元格行的校验信息
            //这一步至关重要这是目前找到的剔除无效行的最好的办法了,当然了我们需要实现IExcelModel接口来实现这一步
            this.errorMsg = "";
        } else {
            this.errorMsg = errorMsg;
        }
    }

    @Override
    public ExcelVerifyHandlerResult verifyHandler(People people) {
        if ("人走茶凉".equals(people.getName())) {
            this.errorMsg = "姓名不符合要求,不能是人走茶凉";
        }
        return null;
    }
}