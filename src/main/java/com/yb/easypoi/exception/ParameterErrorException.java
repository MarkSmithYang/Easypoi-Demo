package com.yb.easypoi.exception;

/**
 * @author yangbiao
 * @Description:常用的自定义运行时异常
 * @date 2018/10/31
 */
public class ParameterErrorException extends RuntimeException{
    private static final long serialVersionUID = 7070951253104458553L;

    public ParameterErrorException(String message) {
        super(message);
    }

    //抛出自己需要的异常信息
    public static void message(String message){
        throw new ParameterErrorException(message);
    }
}
