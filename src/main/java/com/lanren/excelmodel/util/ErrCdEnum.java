package com.lanren.excelmodel.util;

public enum ErrCdEnum {
    C00990001("00990001", "参数为空"),
    C00990002("00990002", "格式错误"),
    C00990003("00990003", "数据不匹配"),
    C00990004("00990004", "信息不存在"),
    C00990005("00990005", "信息已存在"),
    C00990006("00990006", "信息量过大或超长"),
    C00990007("00990007", "SQL执行失败"),
    C00990008("00990008", "程序异常"),
    C00990009("00990009", "规则不符合"),
    C00990010("00990010", "第三方服务异常"),
    C00990011("00990011", "事务回滚"),
    C00990012("00990012", "数据不完整"),
    C00990013("00990013", "分布式锁已开启"),
    C00990099("00990099", "code未定义");

    private String code;
    private String msg;

    ErrCdEnum(String code, String msg) {
        this.code = code;
        this.msg = msg;
    }

    public String getCode() {
        return this.code;
    }

    public String getMsg() {
        return this.msg;
    }
}
