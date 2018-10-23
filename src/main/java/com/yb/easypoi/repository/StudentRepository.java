package com.yb.easypoi.repository;

import com.google.common.collect.Lists;
import com.yb.easypoi.model.Student;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.collections4.MapUtils;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.jdbc.core.JdbcTemplate;
import org.springframework.jdbc.core.namedparam.NamedParameterJdbcTemplate;
import org.springframework.stereotype.Repository;

import java.sql.Date;
import java.time.LocalDate;
import java.util.List;
import java.util.Map;
import java.util.Set;

/**
 * @author yangbiao
 * @Description:持久层代码
 * @date 2018/10/23
 */
@Repository
public class StudentRepository {

    @Autowired
    private JdbcTemplate jdbcTemplate;
    @Autowired
    private NamedParameterJdbcTemplate namedParameterJdbcTemplate;

    /**
     * 查询所有返回list集合装Student
     * @return
     */
    public List<Student> findAll() {
        List<Map<String, Object>> maps = jdbcTemplate.queryForList("select * from student");
        //封装数据
        List<Student> result = Lists.newArrayList();
        if (CollectionUtils.isNotEmpty(maps)) {
            maps.forEach(s -> {
                Student student = new Student();
                student.setId((String) s.get("id"));
                student.setName((String) s.get("name"));
                student.setAge((Integer)s.get("age"));
                student.setClassName((String)s.get("class_name"));
                //强转为sql包下的Date
                student.setJoinTime(((Date) s.get("join_time")).toLocalDate());
                result.add(student);
            });
        }
        return result;
    }

    /**
     * 查询所有返回list集合嵌套map
     * @return
     */
    public List<Map<String, Object>> queryAll() {
        List<Map<String, Object>> result = jdbcTemplate.queryForList("select * from student");
        return result;
    }
}
