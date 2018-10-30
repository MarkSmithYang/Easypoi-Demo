package com.yb.easypoi.service;

import cn.afterturn.easypoi.entity.ImageEntity;
import cn.afterturn.easypoi.excel.ExcelExportUtil;
import cn.afterturn.easypoi.excel.ExcelImportUtil;
import cn.afterturn.easypoi.excel.entity.ExportParams;
import cn.afterturn.easypoi.excel.entity.ImportParams;
import cn.afterturn.easypoi.excel.entity.TemplateExportParams;
import cn.afterturn.easypoi.excel.entity.params.ExcelExportEntity;
import cn.afterturn.easypoi.excel.entity.result.ExcelImportResult;
import com.google.common.collect.Lists;
import com.yb.easypoi.model.Course;
import com.yb.easypoi.model.People;
import com.yb.easypoi.model.Student;
import com.yb.easypoi.model.Teacher;
import com.yb.easypoi.repository.StudentRepository;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

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

    /**
     * 下载Excel文件到指定的位置(一般都是项目resource下)
     *
     * @param workbook
     * @throws IOException
     */
    private void downloadExcel(Workbook workbook) throws IOException {
        //设置文件所在位置的文件夹
        File savefile = new File("src\\main\\resources\\static\\");
        //删除以前的文件,保持文件是最近操作生成的
        if (!savefile.exists()) {
            savefile.mkdirs();
        }
        //获取字节输出流
        FileOutputStream fos = new FileOutputStream("src\\main\\resources\\static\\fail.xls");
        //输出(下载)工作簿
        workbook.write(fos);
        //关闭流
        fos.close();
    }

    /**
     * 用输出的流的方式把Excel在网页输出---(方法抽取)
     *
     * @param response
     * @param word
     */
    public void outPutWord(HttpServletResponse response, XWPFDocument word) {
        if (word == null) {
            log.info("获取到工作簿为空");
            return;
        }
        //以流的方式输出到页面
        //重置响应对象
        response.reset();
        //设置文件名称
        String exportName = "我的word-" + DateTimeFormatter.ofPattern("yyyyMMddHHmmss").format(LocalDateTime.now());
        // 指定下载的文件名--设置响应头
        try {
            response.setHeader("Content-Disposition", "attachment;filename=" +
                    new String(exportName.getBytes("gb2312"), "ISO8859-1") + ".docx");
        } catch (UnsupportedEncodingException e) {
            log.info("中文文件名编码出错");
            e.printStackTrace();
        }
        response.setContentType("application/vnd.ms-word;charset=UTF-8");
        response.setHeader("Pragma", "no-cache");
        response.setHeader("Cache-Control", "no-cache");
        response.setDateHeader("Expires", 0);
        outPutStream(response, null, word);
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
        //设置文件名称
        String exportName = "我的Excel-" + DateTimeFormatter.ofPattern("yyyyMMddHHmmss").format(LocalDateTime.now());
        //重置响应对象
        response.reset();
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
        outPutStream(response, workbook, null);
    }

    /**
     * 输出方法抽取-----(公共方法)
     *
     * @param response
     * @param workbook
     * @param word
     */
    private void outPutStream(HttpServletResponse response, Workbook workbook, XWPFDocument word) {
        //获取文件的流
        try {
            OutputStream outputStream = response.getOutputStream();
            //创建缓冲流
            BufferedOutputStream stream = new BufferedOutputStream(outputStream);
            //用工作簿输出流
            if (word != null) {
                word.write(stream);
            } else if (workbook != null) {
                workbook.write(stream);
            } else {
                throw new IOException("文档参数不能为空");
            }
            //刷新关闭响应的流,先缓冲流
            stream.flush();
            stream.close();
            outputStream.close();
        } catch (IOException e) {
            log.info("response获取数据输出流失败");
            e.printStackTrace();
        }
    }

    /**
     * 查询所有的student数据信息---List<Student></>
     *
     * @return
     */
    public List<Student> findAll() {
        List<Student> result = studentRepository.findAll();
        return result;
    }

    /**
     * 查询所有的student数据信息---List<Map<String,Object></>></>
     *
     * @return
     */
    public List<Map<String, Object>> queryAll() {
        List<Map<String, Object>> result = studentRepository.queryAll();
        return result;
    }

    /**
     * 导出数据到Excel--没用模板的情况
     */
    public void exportFile(HttpServletResponse response) {
        //获取需要导出的数据
        List<Student> all = studentRepository.findAll();
        //设置相关的样式标题等信息
        ExportParams exportParams = new ExportParams();
//        exportParams.setTitle("学生信息");
        //获取工作簿
        Workbook workbook = ExcelExportUtil.exportExcel(exportParams, Student.class, all);
        //输出为Exce文件
        outPutExcel(response, workbook);
    }

    /**
     * 导出数据到Excel--用模板的情况
     */
    public void exportTemplae(HttpServletResponse response) {
        TemplateExportParams params = new TemplateExportParams(
                "src\\main\\resources\\templates\\WPS创建的Excel模板.xlsx");
        params.setSheetName("学生信息表");
        //封装数据
        Map<String, Object> map = new HashMap<String, Object>();
        map.put("id", "入学编号");
        map.put("name", "姓名");
        map.put("age", "年龄");
        map.put("join_time", "入学时间");
        map.put("class_name", "班级");

        //图片导出的写法(直接通过有参构造来设置图片信息,建议使用url的那种,
        // 字节数组的那种是比较特定情况下来用,因为比较繁琐)
        ImageEntity imageEntity1 = new ImageEntity("src/main/resources/static/a.png", 400, 300);
        //map.put("class_name", imageEntity1);
        //查询student的数据
        List<Student> all = studentRepository.findAll();
        //放入数据(需要通过遍历获取数据)
        map.put("mapList", all);
        //获取工作簿,在网页输出
        Workbook workbook = ExcelExportUtil.exportExcel(params, map);
        outPutExcel(response, workbook);
    }

    /**
     * 多个关联对象情况下的Excel导出
     *
     * @param response
     */
    public void exportCollect(HttpServletResponse response) {
        //查询信息
        List<Map<String, Object>> lists = studentRepository.queryCourseList();
        //实例化list集合
        List<Course> result = new ArrayList<>();
        //处理逻辑
        if (CollectionUtils.isNotEmpty(lists)) {
            lists.forEach(s -> {
                Course course = new Course();
                course.setId((String) s.get("id"));
                course.setCourseName((String) s.get("course_name"));
                //查询教师信息
                String teacherId = (String) s.get("teacher_id");
                Teacher teacher = studentRepository.queryTeahcer(teacherId);
                course.setTeacher(teacher);
                //获取学生id
                List<String> ids = studentRepository.queryStudentId(course.getCourseName());
                //查询学生信息
                List<Student> students = studentRepository.queryStudentList(ids);
                course.setList(students);
                result.add(course);
            });
        }
        //导出Excel文件
        ExportParams exportParams = new ExportParams("学校课程", "sheet@1");
        Workbook workbook = ExcelExportUtil.exportExcel(exportParams, Course.class, result);
        outPutExcel(response, workbook);
    }

    /**
     * 通过ExcelExportEntity(一个对象相当于一个注解,比注解更灵活,
     * 注解只能提前写死,而实体这种可以加判断)导出导出Excel数据
     *
     * @param response
     */
    public void exportEntity(HttpServletResponse response) {
        List<ExcelExportEntity> list = Lists.newArrayList();
        //设置id信息
        list.add(new ExcelExportEntity("学号", "id"));
        //设置name信息
        ExcelExportEntity name = new ExcelExportEntity("姓名", "name");
        name.setNeedMerge(true);
        name.setMergeVertical(true);
        list.add(name);
        //设置name信息
        ExcelExportEntity age = new ExcelExportEntity("年龄", "age");
        age.setNeedMerge(true);
        age.setMergeVertical(true);
        list.add(age);
        //设置name信息
        ExcelExportEntity className = new ExcelExportEntity("班级", "class_name");
        className.setNeedMerge(true);
        className.setMergeVertical(true);
        list.add(className);
        //设置name信息
        ExcelExportEntity joinTime = new ExcelExportEntity("入学日期", "join_time");
        joinTime.setNeedMerge(true);
        joinTime.setMergeVertical(true);
        list.add(joinTime);
        //查询导出数据
        List<Map<String, Object>> maps = studentRepository.queryAll();
        //导出数据
        Workbook workbook = ExcelExportUtil.exportExcel(new ExportParams("ExcelExportEntity版的导出", "sheet-1"),
                list, maps);
        //输出到页面
        outPutExcel(response, workbook);
    }


    //000000000000000000000000000000000000000000以下是导入的内容0000000000000000000000000000000000000000000000

    /**
     * 通过Map这种方式的导入
     * * @return
     *
     * @param file
     * @param response
     */
    public String importFile(MultipartFile file, HttpServletResponse response) {
        if (file == null) {
            log.info("上传附件为空");
            return "导入失败";
        }
        //设置导入的一些基本信息
        ImportParams importParams = new ImportParams();
        //如果有标题必须要正确的设置标题行数(默认0行-->没标题)
        importParams.setTitleRows(1);
        //表头默认是一行,如果不是请正确设置
        importParams.setHeadRows(1);
        //(数据)开始行默认是0,也就是表头下的第一行数据是0,一般不跳过前几行数据也不用设置
        importParams.setStartRows(0);
        //开启导入校验
        importParams.setNeedVerfiy(true);
        try {
            //一般来说都是上传附件,故而基本都是输入流的方式
            InputStream iss = file.getInputStream();
            List<Map<String, Object>> maps = ExcelImportUtil.importExcel(iss, Map.class, importParams);
            //流被使用了就用完了,所以需要在次获取流(实测)
            InputStream is = file.getInputStream();
            //使用这个api可以获取导入校验报错的信息
            ExcelImportResult<People> result = ExcelImportUtil.importExcelMore(is, People.class, importParams);
            //获取导入成功的数据
            List<People> list = result.getList();
            //获取导入错误的数据
            List<People> failList = result.getFailList();
            //是否校验错误
            boolean verfiyFail = result.isVerfiyFail();
            System.err.println(failList);
            System.err.println("==" + verfiyFail);
            System.err.println(list);
            System.err.println(maps);
            //获取错误的工作簿
            Workbook failWorkbook = result.getFailWorkbook();
            //把错误的工作簿信息存储到指定的位置,以便下载阅看
            downloadExcel(failWorkbook);
            //保存数据到数据库
            if(!result.isVerfiyFail()){
                System.err.println("导入未出现错误");
            }else{
                System.err.println("导入错误");
            }
        } catch (Exception e) {
            log.info("异常为==" + e.getMessage());
            e.printStackTrace();
            return "导入失败";
        }
        return "导入成功";
    }

}
