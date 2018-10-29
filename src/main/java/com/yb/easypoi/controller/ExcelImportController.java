package com.yb.easypoi.controller;

import com.yb.easypoi.service.StudentService;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.CrossOrigin;
import org.springframework.web.bind.annotation.RestController;

/**
 * @author yangbiao
 * @Description:控制层服务--Excel导入
 * @date 2018/10/29
 */
@RestController
@CrossOrigin//处理跨域访问--(只有同一个域名下的不同服务那种才不是跨域访问)
public class ExcelImportController {
 public static final Logger log = LoggerFactory.getLogger(ExcelImportController.class);

    @Autowired
    private StudentService studentService;


    
}
