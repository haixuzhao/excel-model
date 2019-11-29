package com.lanren.excelmodel.poi;
import com.alibaba.excel.annotation.ExcelProperty;

import java.util.Date;

/**
 * 基础数据类.这里的排序和excel里面的排序一致
 **/
public class DemoData {
    @ExcelProperty("姓名")
    private String string;
    @ExcelProperty("下单日期")
    private Date date;
    @ExcelProperty("金额")
    private Double doubleData;

    public String getString() {
        return string;
    }

    public void setString(String string) {
        this.string = string;
    }

    public Date getDate() {
        return date;
    }

    public void setDate(Date date) {
        this.date = date;
    }

    public Double getDoubleData() {
        return doubleData;
    }

    public void setDoubleData(Double doubleData) {
        this.doubleData = doubleData;
    }
}
