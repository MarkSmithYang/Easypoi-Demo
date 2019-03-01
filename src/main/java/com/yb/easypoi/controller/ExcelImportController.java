package com.yb.easypoi.controller;

import com.yb.easypoi.service.StudentService;
import org.apache.poi.ss.usermodel.Workbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.io.*;

/**
 * @author yangbiao
 * @Description:控制层服务--Excel导入
 * @date 2018/10/29
 */
@Controller
@CrossOrigin//处理跨域访问--(只有同一个域名下的不同服务那种才不是跨域访问)
public class ExcelImportController {
    public static final Logger log = LoggerFactory.getLogger(ExcelImportController.class);

    @Autowired
    private StudentService studentService;

    /**
     * 注意这个资源需要在static下,实测其他的没法访问
     * 类注解必须使用@Controller才行,注意浏览器的缓存问题
     * @return
     */
    @GetMapping("importFailExcel")
    public String importFailExcel() {
      return "mergeFailList.xls";
    }

    /**
     * 通过Map这种方式的导入
     *
     * @param file
     * @return
     */
    @PostMapping("importFile")//导入注意日期时间,这个最容易出现问题,使用java8的LocalDate比较容易出问题,而Date却不会
    @ResponseBody
    public String importFile(MultipartFile file) {
        //使用json形式的数据返回在类上有@Controller的时候需要在方法上添加@ResponseBody
        String result = studentService.importFile(file);
        return result;
    }


}
