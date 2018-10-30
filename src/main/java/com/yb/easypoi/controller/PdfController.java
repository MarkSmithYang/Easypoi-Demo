package com.yb.easypoi.controller;

import cn.afterturn.easypoi.pdf.PdfExportUtil;
import cn.afterturn.easypoi.pdf.entity.PdfExportParams;
import com.yb.easypoi.model.Student;
import com.yb.easypoi.service.StudentService;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.CrossOrigin;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.util.List;

/**
 * @author yangbiao
 * @Description:控制层接口
 * @date 2018/10/29
 */
@RestController
@CrossOrigin//处理跨域访问--(只有同一个域名下的不同服务那种才不是跨域访问)
public class PdfController {
    public static final Logger log = LoggerFactory.getLogger(PdfController.class);

    @Autowired
    private StudentService studentService;

    @GetMapping("exportPdf")
    public void exportPdf(HttpServletResponse response) {
        //设置导出参数
        PdfExportParams pdfExportParams = new PdfExportParams("我是主标题", "我是副标题");
        //查询基础数据
        List<Student> all = studentService.findAll();
        try {
            //目前还不知道为何不支持那个字体,实测中文没法显示,需要添加itext-asian的5.2.0及其以上版本解决
            //com.itextpdf.text.DocumentException: Font 'STSong-Light' with 'UniGB-UCS2-H' is not recognized.
            PdfExportUtil.exportPdf(pdfExportParams, Student.class, all, response.getOutputStream());
        } catch (IOException e) {
            log.info("PDF导出出错==" + e.getMessage());
            e.printStackTrace();
        }
    }

}
