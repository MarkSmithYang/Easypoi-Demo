package com.yb.easypoi.repository;

import com.google.common.collect.Lists;
import com.yb.easypoi.exception.ParameterErrorException;
import com.yb.easypoi.model.People;
import com.yb.easypoi.model.Student;
import com.yb.easypoi.model.Teacher;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.lang3.StringUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.jdbc.core.JdbcTemplate;
import org.springframework.jdbc.core.namedparam.NamedParameterJdbcTemplate;
import org.springframework.stereotype.Repository;
import org.springframework.transaction.annotation.Transactional;

import java.sql.Date;
import java.util.*;
import java.util.stream.Collectors;

/**
 * @author yangbiao
 * @Description:持久层代码
 * @date 2018/10/23
 */
@Repository
public class StudentRepository {
 public static final Logger log = LoggerFactory.getLogger(StudentRepository.class);

    @Autowired
    private JdbcTemplate jdbcTemplate;
    @Autowired
    private NamedParameterJdbcTemplate namedParameterJdbcTemplate;

    /**
     * 把List<Map<String, Object>>转换为List<Student></>
     *
     * @param maps
     * @return
     */
    public List<Student> getStudentList(List<Map<String, Object>> maps) {
        //实例化集合
        List<Student> result = Lists.newArrayList();
        //处理逻辑
        if (CollectionUtils.isNotEmpty(maps)) {
            maps.forEach(s -> {
                Student student = new Student();
                student.setId((String) s.get("id"));
                student.setName((String) s.get("name"));
                student.setAge((Integer) s.get("age"));
                student.setClassName((String) s.get("class_name"));
                //强转为sql包下的Date
                student.setJoinTime(((Date) s.get("join_time")).toLocalDate());
                result.add(student);
            });
        }
        return result;
    }

    /**
     * 查询所有返回list集合装Student
     *
     * @return
     */
    public List<Student> findAll() {
        List<Map<String, Object>> maps = jdbcTemplate.queryForList("select * from student");
        //封装数据返回
        return getStudentList(maps);
    }

    /**
     * 查询所有返回list集合嵌套map---student
     *
     * @return
     */
    public List<Map<String, Object>> queryAll() {
        List<Map<String, Object>> result = jdbcTemplate.queryForList("select * from student");
        return result;
    }

    /**
     * 查询所有课程Course返回list----course
     *
     * @return
     */
    public List<Map<String, Object>> queryCourseList() {
        List<Map<String, Object>> result = jdbcTemplate.queryForList("SELECT * FROM course group by course_name");
        return result;
    }

    /**
     * 根据教师teacher_id查询教师信息
     * 这里设定的是教师和课程是一对一1:1)
     *
     * @return
     */
    public Teacher queryTeahcer(String teacherId) {
        //判断id是否为空
        if (StringUtils.isBlank(teacherId)) {
            return null;
        }
        //实例化一个list集合
        List<Teacher> list = new ArrayList<>();
        //查询数据
        List<Map<String, Object>> result = jdbcTemplate.queryForList("SELECT * FROM teacher where id=?", teacherId);
        if (CollectionUtils.isNotEmpty(result)) {
            result.forEach(s -> {
                Teacher teacher = new Teacher();
                teacher.setId((String) s.get("id"));
                teacher.setName((String) s.get("name"));
                list.add(teacher);
            });
        }
        return CollectionUtils.isNotEmpty(list) ? list.get(0) : null;
    }

    /**
     * 根据学生student_id查询学生信息
     *
     * @return
     */
    public List<Student> queryStudentList(List<String> ids) {
        //判断id是否为空
        if (CollectionUtils.isEmpty(ids)) {
            return null;
        }
        //实例化一个map集合
        Map<String, Object> params = new HashMap<>();
        params.put("ids", ids);
        //查询数据
        List<Map<String, Object>> result = namedParameterJdbcTemplate.
                queryForList("SELECT * FROM student where id in(:ids)", params);
        return getStudentList(result);
    }

    public List<String> queryStudentId(String courseName) {
        //判断id是否为空
        if (StringUtils.isBlank(courseName)) {
            return null;
        }
        //查询数据
        List<Map<String, Object>> result = jdbcTemplate.
                queryForList("SELECT student_id FROM course where course_name = ?", courseName);
        if (CollectionUtils.isNotEmpty(result)) {
            List<String> ids = result.stream().map(s -> {
                return (String) s.get("student_id");
            }).collect(Collectors.toList());
            return ids;
        }
        return null;
    }

    /**
     * 保存导入的数据到数据库
     * @param peoples
     */
    @Transactional
    public void insertPeople(List<People> peoples) {
        if(CollectionUtils.isEmpty(peoples)){
            log.info("获取到的导入数据为空");
            ParameterErrorException.message("导入失败");
        }
        //因为传递过来的集合之前已经判定不为空了所以直接用即可
        peoples.forEach(s->{
            jdbcTemplate.update("insert into people (id,`name`,address) values (?,?,?)",preparedStatement -> {
                preparedStatement.setObject(1, s.getId());
                preparedStatement.setObject(2, s.getName());
                preparedStatement.setObject(3, s.getAddress());
            });
        });
    }
}
