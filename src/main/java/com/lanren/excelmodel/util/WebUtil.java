package com.lanren.excelmodel.util;

import com.alibaba.fastjson.JSONObject;

/**
 * @ClassName WebUtil
 * @Description: TODO
 * @Author zhx
 * @Date 2019/11/28
 * @Version V1.0
 **/
public class WebUtil {
    public static String successResp(Object data, String msg) {
        JSONObject retJson = new JSONObject();
        retJson.put("success", true);
        retJson.put("code", "OK");
        retJson.put("msg", msg == null ? "" : msg);
        retJson.put("data", data == null ? new JSONObject() : data);
        retJson.put("icon", "1");
        return retJson.toJSONString();
    }
    /**
     * 组装 错误/异常 JSON
     * @param code
     * @param msg
     * @return json
     */
    public static String errorResp(String code, String msg, Object data) {
        JSONObject retJson = new JSONObject();
        retJson.put("success", false);
        retJson.put("code", code);
        retJson.put("msg", msg);
        retJson.put("data", data == null ? new JSONObject() : data);
        retJson.put("icon", "2");
        return retJson.toString();
    }
}
