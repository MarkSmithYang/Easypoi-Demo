package com.yb.easypoi.controller;

import cn.afterturn.easypoi.cache.manager.POICacheManager;
import cn.afterturn.easypoi.excel.ExcelXorHtmlUtil;
import cn.afterturn.easypoi.excel.entity.ExcelToHtmlParams;
import cn.afterturn.easypoi.excel.entity.enmus.ExcelType;
import cn.afterturn.easypoi.excel.entity.params.ExcelExportEntity;
import com.yb.easypoi.service.StudentService;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.validation.annotation.Validated;
import org.springframework.web.bind.annotation.CrossOrigin;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.InputStream;
import java.util.Scanner;

/**
 * @author yangbiao
 * @Description:控制层接口
 * @date 2018/10/23
 */
@RestController
@CrossOrigin//处理跨域访问--(只有同一个域名下的不同服务那种才不是跨域访问)
@Validated
public class ExcelExportController {
    public static final Logger log = LoggerFactory.getLogger(ExcelExportController.class);

    @Autowired
    private StudentService studentService;

    /**
     * 通过注解导出Excel数据
     * @param response
     */
    @GetMapping("exportFile")
    public void exportFile(HttpServletResponse response) {
        studentService.exportFile(response);
    }

    /**
     * 通过ExcelExportEntity(一个对象相当于一个注解,比注解更灵活,
     * 注解只能提前写死,而实体这种可以加判断)导出导出Excel数据
     * @param response
     */
    @GetMapping("exportEntity")
    public void exportEntity(HttpServletResponse response) {
        studentService.exportEntity(response);
    }

    //--------------------------------------------------------------------------------------------

    /**
     * 通过模板导出Excel数据
     * @param response
     */
    @GetMapping("exportTemplae")
    public void exportTemplae(HttpServletResponse response) {
        studentService.exportTemplae(response);
    }

    //--------------------------------------------------------------------------------------------

    /**
     * 多个关联对象情况下的Excel导出--->设定教师和课程是一对一的,因为是用jdbc做的,
     * 数据不好封装,所以有很多地方会很繁琐,使用easypoi的时候用jpa是最好的
     * @param response
     */
    @GetMapping("exportCollect")
    public void exportCollect(HttpServletResponse response) {
        studentService.exportCollect(response);
    }

    //--------------------------------------------------------------------------------------------

    /**
     * 把Excel响应到网页预览
     * @param response
     */
    @GetMapping("excelToHtml")
    public void excelToHtml(HttpServletResponse response) {
        //直接获取byte数组的数据,然后response响应到页面即可
        try {
            ExcelToHtmlParams params = new ExcelToHtmlParams(WorkbookFactory.create(POICacheManager.
                    getFile("src\\main\\resources\\templates\\WPS创建的Excel模板.xlsx")));
            response.getOutputStream().write(ExcelXorHtmlUtil.excelToHtml(params).getBytes());
        } catch (IOException e) {
            log.info("异常信息==" + e.getMessage());
            e.printStackTrace();
        }
    }

    //------------------------------------以下是源码有个类找不到------------------------------------------

    /**
     * 字符串形式的Html转为Excel---(推荐)--(源码有bug,源码的有个类找不到)
     * @throws Exception
     */
    @GetMapping("htmlToExcelByIs")
    public void htmlToExcelByIs(HttpServletResponse response) {
        Workbook workbook = ExcelXorHtmlUtil.htmlToExcel(getClass().
                getResourceAsStream("/static/sample.html"), ExcelType.XSSF);
        studentService.outPutExcel(response, workbook);
    }

    /**
     * 字符串形式的Html转为Excel--(源码有bug,源码的有个类找不到)
     * @throws Exception
     */
    @GetMapping("htmlToExcelByStr")
    public void htmlToExcelByStr(HttpServletResponse response) {
        StringBuilder html = new StringBuilder();
        InputStream inputStream = getClass().
                getResourceAsStream("/static/sample.html");
        Scanner s = new Scanner(inputStream, "utf-8");
        while (s.hasNext()) {
            html.append(s.nextLine());
        }
        s.close();
        //获取workbook的时候,源码里面的Jsoup找不到,java.lang.ClassNotFoundException: org.jsoup.Jsoup
        Workbook workbook = ExcelXorHtmlUtil.htmlToExcel(html.toString(), ExcelType.XSSF);
        studentService.outPutExcel(response, workbook);
    }
}
