package com.lanren.excelmodel.util;

import com.alibaba.fastjson.JSONObject;
import com.lanren.excelmodel.poi.DemoData;

import java.util.ArrayList;
import java.util.Date;
import java.util.List;

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


    public static List<String> getTitles(){
        List<String> retList = new ArrayList<>();
        retList.add("姓名");
        retList.add("下单日期");
        retList.add("金额");
        return retList;
    }

    public static List<DemoData> data() {
        List<DemoData> list = new ArrayList<DemoData>();
        for (int i = 0; i < 10; i++) {
            DemoData data = new DemoData();
            data.setString("字符串" + i);
            data.setDate(new Date());
            data.setDoubleData(0.56);
            list.add(data);
        }
        return list;
    }

    public static List<List<String>> dataList() {
        List<List<String>> list = new ArrayList<>();
        for (int i = 0; i < 10; i++) {
            List<String> data = new ArrayList<>();
            data.add("字符串" + i);
            data.add("2019-09-09 00:00:00");
            data.add("0.56");
            list.add(data);
        }
        return list;
    }
}
