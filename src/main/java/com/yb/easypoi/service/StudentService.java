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
import com.yb.easypoi.exception.ParameterErrorException;
import com.yb.easypoi.model.Course;
import com.yb.easypoi.model.People;
import com.yb.easypoi.model.Student;
import com.yb.easypoi.model.Teacher;
import com.yb.easypoi.repository.StudentRepository;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.stream.Collectors;

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
    private void downloadExcel(Workbook workbook, String dirPath, String fileName) throws IOException {
        if (workbook == null) {
            log.info("传入的工作簿为空");
            ParameterErrorException.message("操作失败");
        }
        //判断路径是否靠谱
        if (StringUtils.isBlank(dirPath) || StringUtils.isBlank(fileName)) {
            log.info("文件夹路径或文件名为空");
            ParameterErrorException.message("操作失败");
        }
        //设置文件所在位置的文件夹
        File savefile = new File(dirPath);
        //删除以前的文件,保持文件是最近操作生成的
        if (!savefile.exists()) {
            savefile.mkdirs();
        }
        //获取字节输出流
        FileOutputStream fos = new FileOutputStream(dirPath + fileName);
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
        exportParams.setTitle("学生信息");
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


    //000000000000000000000000000000000000000000以下是导入的内容000000000000000000000000000000000000000000000

    /**
     * 通过Map这种方式的导入
     * * @return
     *
     * @param file
     * @param response
     */
    @Transactional(rollbackFor = Exception.class)
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
            //-------------------------------------------------------------------------------------------
            //一般来说都是上传附件,故而基本都是输入流的方式
            InputStream iss = file.getInputStream();
            List<Map<String, Object>> maps = ExcelImportUtil.importExcel(iss, Map.class, importParams);
            System.err.println("通过map封装的数据:" + maps);
            //流被使用了就用完了,所以需要在次获取流(实测)
            InputStream isss = file.getInputStream();
            //这个是直接获取集合的api,这个获取不到fail相关的数据,但是自定义校验依旧可以得到-->不推荐使用这个
            List<People> objects = ExcelImportUtil.importExcel(isss, People.class, importParams);
            //-------------------------------------------------------------------------------------------

            //流被使用了就用完了,所以需要在次获取流(实测)
            InputStream is = file.getInputStream();
            //使用这个api可以获取导入校验报错的信息
            ExcelImportResult<People> result = ExcelImportUtil.importExcelMore(is, People.class, importParams);
            if (result == null) {
                log.info("返回的Excel的导入结果对象为空");
                ParameterErrorException.message("导入错误");
            }
            //获取整合后校验信息的集合
            List<People> collect = getMergeFailWordbook(result);
            //获取getList()的数据-->可能包含自定义校验不通过的信息但是它是注解(常规的校验(非自定义的))校验通过的
            List<People> list = result.getList();
            System.err.println(list);
            //获取导入错误的数据
            List<People> failList = result.getFailList();
            System.err.println(failList);
            //制作一个开关按钮-->用于处理含有无效行的时候result.isVerfiyFail()为false的情况
            boolean flag = false;
            //实测证明不管result.isVerfiyFail()为true还是false,都不会影响到自定义校验的,自定义校验未通过时,
            // 这个还是可以false,只有当注解校验不通过时,它才会为true,所以这个不能作为是否校验通过的的依据
            //当然了,没有自定义校验的时候,这个应该可以用,但是为了兼容性,就不会去用它
            boolean verfiyFail = result.isVerfiyFail();
            //获取错误的工作簿-->实测证明自定义校验不通过的也不会再里面
            Workbook failWorkbook = result.getFailWorkbook();
            //把响应的信息存储在项目下以便查看或下载
            downloadExcel(failWorkbook, "src\\main\\resources\\static\\", "failWorkbook.xls");
            List<String> errors = new ArrayList<>();
            //以json的形式打印出校验不通过的信息--->可以根据自己的情况来定,我这里没有做返回值,而是做成Excel
            if (CollectionUtils.isNotEmpty(collect)) {
                //校验信息按照行号正序排序
                collect.sort((a, b) -> {
                    if (a.getRowNum() >= b.getRowNum()) {
                        return 1;
                    } else {
                        return -1;
                    }
                });
                //提取校验错误的信息
                collect.forEach(s -> {
                    errors.add("第" + s.getRowNum() + "行:" + s.getErrorMsg());
                });
                log.info("导入错误信息:" + errors);
                //获取合并后的工作簿
                Workbook mergeFail = ExcelExportUtil.exportExcel(new ExportParams("错误数据的校验情况", "sheet@@@"), People.class, collect);
                //存储Excel
                downloadExcel(mergeFail, "src\\main\\resources\\static\\", "mergeFailList.xls");
                return "导入错误,错误信息为:" + errors;
            } else {
                //没有有效的校验信息,即没有导入错误或者传入的 ExcelImportResult<People>对象为空
                flag = true;
            }
            //保存数据
            if (flag) {
                //导入工作簿没有出现错误,处理正确的数据保存数据到数据库
                if (CollectionUtils.isNotEmpty(list)) {
                    //因为上面合并校验的时候没有出现校验不通过的信息,说明list里已经没有自定义的校验不通过的信息存在
                    //所以这里的数据全部都是导入成功无异常的数据-->所以可以直接保存数据
                    //实测证明,当没有校验的时候,无效行也会随着进入到getList()数据中,如果想保存到数据库中,需要剔除它们
                    List<People> peoples = list.stream().filter(s -> {
                        //当封装数据的对象需要封装的属性的值都为空时,说明是无效对象,剔除它们
                        if (StringUtils.isBlank(s.getId()) && StringUtils.isBlank(s.getName()) &&
                                StringUtils.isBlank(s.getAddress())) {
                            return false;
                        } else {
                            return true;
                        }
                    }).collect(Collectors.toList());
                    //保存数据
                    studentRepository.insertPeople(peoples);
                    return "导入成功";
                } else {
                    log.info("导入的数据为空");
                }
            }
        } catch (Exception e) {
            log.info("异常为==" + e.getMessage());
            e.printStackTrace();
            //抛出异常让事务回滚-->重点
            ParameterErrorException.message(e.getMessage());
        }
        return "导入失败";
    }

    /**
     * 合并整合校验不通过的校验信息
     *
     * @param result
     * @return
     */
    private List<People> getMergeFailWordbook(ExcelImportResult<People> result) {
        //初始化一个集合备用
        List<People> people = new ArrayList<>();
        //获取整合后校验不通过的数据
        //获取list和failList的数据
        List<People> li = result.getList();
        List<People> fa = result.getFailList();
        //合并两集合的内容
        if (CollectionUtils.isNotEmpty(li)) {
            //获取有效的校验信息(就是通过自定义校验不通过的信息-->实测它会进入这个list里和正常的数据在一起)
            List<People> collect = li.stream().filter(s -> {
                //剔除校验通过的信息
                if (StringUtils.isNotBlank(s.getErrorMsg())) {
                    return true;
                } else {
                    return false;
                }
            }).collect(Collectors.toList());
            //实测证明集合不能合并null--->即list.addAll(null)是会报空指针的
            //也不能直接通过people来接收,万一为null,那么people就白实例化了下面使用肯定会报错的
            if (CollectionUtils.isNotEmpty(collect)) {
                people.addAll(collect);
            }
        }
        //如果failList不为空则合并
        if (CollectionUtils.isNotEmpty(fa)) {
            //剔除无效的信息
            List<People> fails = fa.stream().filter(a -> {
                //通过自己定义的给无效行定义的标记,把无效行剔除
                // 我这里没有弄常量字典->标记尽量不太可能和一般的校验信息一样,不然有些有效行也会被剔除
                //如果还是怕的话,就用UUID生成一个字符串来标记吧,我相信能够重合的概率几乎不存在了
                if ("^invalid&^@error^".equals(a.getErrorMsg())) {
                    return false;
                } else {
                    return true;
                }
            }).collect(Collectors.toList());
            //剔除无效行之后,如果不为空,说明有有效的校验信息,合并到集合里
            if (CollectionUtils.isNotEmpty(fails)) {
                people.addAll(fails);
            }
        }
        return people;
    }
}
