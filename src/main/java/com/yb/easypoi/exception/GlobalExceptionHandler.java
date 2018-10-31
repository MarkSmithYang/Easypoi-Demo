package com.yb.easypoi.exception;

import com.alibaba.fastjson.JSONObject;
import com.fasterxml.jackson.databind.util.JSONPObject;
import org.springframework.http.HttpStatus;
import org.springframework.web.bind.annotation.ExceptionHandler;
import org.springframework.web.bind.annotation.ResponseStatus;
import org.springframework.web.bind.annotation.RestControllerAdvice;

/**
 * @author yangbiao
 * @Description:controller层异常统一处理类
 * @date 2018/10/31
 */
@RestControllerAdvice//同样的以Rest开头的这个注解,会自动处理成json的形式,不用再在异常方法上写@ResponseBody了
public class GlobalExceptionHandler {

    @ResponseStatus(HttpStatus.BAD_REQUEST)
    @ExceptionHandler(ParameterErrorException.class)
    public JSONObject parameterErrorExceptionHandler(ParameterErrorException e) {
        JSONObject jsonObject = new JSONObject();
        jsonObject.put("status", HttpStatus.BAD_REQUEST.value());
        jsonObject.put("message", e.getMessage());
        return jsonObject;
    }

}
