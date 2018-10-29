package com.yb.easypoi.controller;

import cn.afterturn.easypoi.entity.ImageEntity;
import cn.afterturn.easypoi.word.WordExportUtil;
import com.yb.easypoi.service.StudentService;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.CrossOrigin;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;
import javax.servlet.http.HttpServletResponse;
import java.time.LocalDate;
import java.util.HashMap;
import java.util.Map;

/**
 * @author yangbiao
 * @Description:控制层接口
 * @date 2018/10/26
 */
@RestController
@CrossOrigin//处理跨域访问--(只有同一个域名下的不同服务那种才不是跨域访问)
public class WordController {
    public static final Logger log = LoggerFactory.getLogger(WordController.class);

    @Autowired
    private StudentService studentService;

    @GetMapping("exportWord")
    public void exportWord(HttpServletResponse response) {
        //填充数据
        Map<String, Object> map = new HashMap<>();
        map.put("department", "Easypoi");
        map.put("p.name", "JueYue");
        map.put("time", LocalDate.now());
        //如果有图片,设置相应的图片信息
        ImageEntity imageEntity = new ImageEntity("src/main/resources/static/a.png", 400, 300);
        //封装进map
        map.put("testCode", imageEntity);
        //导出数据为word
        try {
            //传入word的模板,里面和Excel的表达式一样
            XWPFDocument word07 = WordExportUtil.exportWord07("src\\main\\resources\\static\\word.docx", map);
            studentService.outPutWord(response, word07);
        } catch (Exception e) {
            log.info("异常为==" + e.getMessage());
            e.printStackTrace();
        }
    }

}
