package com.yb.easypoi.controller;

import cn.afterturn.easypoi.cache.manager.POICacheManager;
import cn.afterturn.easypoi.excel.ExcelExportUtil;
import cn.afterturn.easypoi.excel.ExcelImportUtil;
import cn.afterturn.easypoi.excel.ExcelXorHtmlUtil;
import cn.afterturn.easypoi.excel.entity.ExcelToHtmlParams;
import cn.afterturn.easypoi.excel.entity.ImportParams;
import cn.afterturn.easypoi.excel.entity.TemplateExportParams;
import com.yb.easypoi.model.Student;
import com.yb.easypoi.service.StudentService;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.CrossOrigin;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

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
    public void exportFile(HttpServletResponse response) {
        studentService.exportFile(response);
    }

    @GetMapping("export")
    public void importFile(HttpServletResponse response) throws Exception {
        studentService.export(response);
    }

    @GetMapping("exportTemplae")
    public void exportTemplae(HttpServletResponse response) throws Exception {
        TemplateExportParams params = new TemplateExportParams(
                "C:\\MyDemoRepository\\easypoi-demo\\src\\main\\resources\\templates\\dd.xls");
        Map<String, Object> map = new HashMap<String, Object>();
        map.put("id", "入学编号");
        map.put("name", "姓名");
        map.put("age", "年龄");
        map.put("class_name", "班级");
        map.put("join_time", "入学时间");

        List<Student> all = studentService.findAll();
        map.put("maplist", all);

        Workbook workbook = ExcelExportUtil.exportExcel(params, map);
        studentService.outPutExcel(response, workbook);
//        File savefile = new File("D:/excel/");
//        if (!savefile.exists()) {
//            savefile.mkdirs();
//        }
//        FileOutputStream fos = new FileOutputStream("D:/excel/专项支出用款申请书_map.xls");
//        workbook.write(fos);
//        fos.close();
    }


}
