package com.yb.easypoi.controller;

import com.yb.easypoi.service.StudentService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.CrossOrigin;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletResponse;

/**
 * @author yangbiao
 * @Description:控制层接口
 * @date 2018/10/23
 */
@RestController
@CrossOrigin//处理跨域访问--(只有同一个域名下的不同服务那种才不是跨域访问)
public class StudentController {

    @Autowired
    private StudentService studentService;

    @GetMapping("exportFile")
    public void exportFile(HttpServletResponse response){
        studentService.exportFile(response);
    }
}
