package com.yb.easypoi.service;

import cn.afterturn.easypoi.cache.manager.POICacheManager;
import cn.afterturn.easypoi.excel.ExcelExportUtil;
import cn.afterturn.easypoi.excel.ExcelXorHtmlUtil;
import cn.afterturn.easypoi.excel.entity.ExcelToHtmlParams;
import cn.afterturn.easypoi.excel.entity.ExportParams;
import cn.afterturn.easypoi.excel.entity.TemplateExportParams;
import cn.afterturn.easypoi.excel.entity.params.ExcelExportEntity;
import cn.afterturn.easypoi.excel.export.template.ExcelExportOfTemplateUtil;
import com.yb.easypoi.model.Student;
import com.yb.easypoi.repository.StudentRepository;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.poi.POITextExtractor;
import org.apache.poi.POIXMLDocument;
import org.apache.poi.extractor.ExtractorFactory;
import org.apache.poi.hssf.model.InternalWorkbook;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.poifs.filesystem.DocumentFactoryHelper;
import org.apache.poi.poifs.filesystem.NPOIFSFileSystem;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.tomcat.util.http.fileupload.FileItemStream;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.*;

/**
 * @author yangbiao
 * @Description:服务层代码
 * @date 2018/10/23
 */
@Service
public class StudentService {
    public static final Logger log = LoggerFactory.getLogger(StudentService.class);

    @Autowired
    private StudentRepository studentRepository;

    public List<Student> findAll(){
        List<Student> result = studentRepository.findAll();
        return result;
    }

    /**
     * 导出数据到Excel--没用模板的情况
     * --注意ExcelExportUtil(3.3.0版本)和ExcelUtils(应该是较低版本的api(3.0.3还是这样的))
     */
    public void exportFile(HttpServletResponse response) {
        //获取需要导出的数据
        List<Student> all = studentRepository.findAll();
        //设置相关的样式标题等信息
        ExportParams exportParams = new ExportParams();
        exportParams.setTitle("学生信息");
        //获取工作簿
        Workbook workbook = ExcelExportUtil.exportExcel(exportParams, Student.class, all);
        //输出为Exce文件
        outPutExcel(response, workbook);
    }

    /**
     * 导出数据到Excel--用模板的情况
     *
     * @param response
     */
    public void export(HttpServletResponse response) {
        try {
            InputStream inputStream = new FileInputStream(new File("C:\\MyDemoRepository\\easypoi-demo\\src\\main\\resources\\templates\\我的模板.xls"));
            TemplateExportParams params = new TemplateExportParams();
            params.setHeadingStartRow(0);
            params.setHeadingRows(2);
            params.setDataSheetNum(3);
            params.setSheetName("学生");
            params.setTemplateUrl("C:\\MyDemoRepository\\easypoi-demo\\src\\main\\resources\\templates\\学生信息模板.xlsx");
            Map<String, Object> map = new HashMap<String, Object>();

            List<Student> all = studentRepository.findAll();
            map.put("list", all);
            Workbook exportExcel = ExcelExportUtil.exportExcel(params, map);
            outPutExcel(response, exportExcel);

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }

    }

    /**
     * 用输出的流的方式把Excel在网页输出---(方法抽取)
     *
     * @param response
     * @param workbook
     */
    public void outPutExcel(HttpServletResponse response, Workbook workbook) {
        if (workbook == null) {
            log.info("获取到工作簿为空");
            return;
        }
        //以流的方式输出到页面
        //重置响应对象
        response.reset();
        //设置文件名称
        String exportName = "我的学生表-" + DateTimeFormatter.ofPattern("yyyyMMddHHmmss").format(LocalDateTime.now());
        // 指定下载的文件名--设置响应头
        try {
            response.setHeader("Content-Disposition", "attachment;filename=" +
                    new String(exportName.getBytes("gb2312"), "ISO8859-1") + ".xls");
        } catch (UnsupportedEncodingException e) {
            log.info("中文文件名编码出错");
            e.printStackTrace();
        }
        response.setContentType("application/vnd.ms-excel;charset=UTF-8");
        response.setHeader("Pragma", "no-cache");
        response.setHeader("Cache-Control", "no-cache");
        response.setDateHeader("Expires", 0);
        //获取文件的流
        try {
            OutputStream outputStream = response.getOutputStream();
            //创建缓冲流
            BufferedOutputStream stream = new BufferedOutputStream(outputStream);
            //用工作簿输出流
            workbook.write(stream);
            //刷新关闭响应的流,先缓冲流
            stream.flush();
            stream.close();
            outputStream.close();
        } catch (IOException e) {
            log.info("response获取数据输出流失败");
            e.printStackTrace();
        }
    }
}
