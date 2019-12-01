package com.lanren.excelmodel.controller;

import com.alibaba.excel.EasyExcel;
import com.alibaba.fastjson.JSON;
import com.lanren.excelmodel.easyexcel.DemoData;
import com.lanren.excelmodel.easyexcel.DemoDataListener;
import com.lanren.excelmodel.poi.ExcelUtil;
import com.lanren.excelmodel.util.ErrCdEnum;
import com.lanren.excelmodel.util.TestFileUtil;
import com.lanren.excelmodel.util.WebUtil;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.URLEncoder;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

/**
 * @ClassName ExcelController
 * @Description: TODO
 * @Author zhx
 * @Date 2019/11/28
 * @Version V1.0
 **/
@Controller
@RequestMapping("excel/*")
public class ExcelController {
    //定义一个全局的记录器，通过LoggerFactory获取
    private final static Logger logger = LoggerFactory.getLogger(ExcelController.class);

    @RequestMapping("index")
    public String toIndexView(Model model){
        return "index";
    }

    @RequestMapping("upload")
    @ResponseBody
    public String upload(@RequestParam("file") MultipartFile multipartFile, @RequestParam Map<String,Object> params){
        try{
            String data = ("10".equals(params.get("toolCd")))?analysisByPOI(multipartFile):analysisByEasy(multipartFile);
            return WebUtil.successResp(data,"导入成功");
        }catch(Exception e){
            e.printStackTrace();
            logger.error("解析失败",e);
            return WebUtil.errorResp(ErrCdEnum.C00990008.getCode(),ErrCdEnum.C00990008.getMsg(),"");
        }
    }

    @RequestMapping("download")
    @ResponseBody
    public String download(HttpServletResponse response){
        ServletOutputStream outputStream = null;
        try {
            response.setContentType("application/binary;charset=UTF-8");
            response.setHeader("Content-Disposition", "attachment;fileName=" + URLEncoder.encode("【dataDemo】" + ".xlsx", "UTF-8"));
            outputStream = response.getOutputStream();
            String fileName = TestFileUtil.getPath() + "file/dataDemo.xlsx";
            File file = new File(fileName);
            IOUtils.copy(new FileInputStream(file),outputStream);
        } catch (Exception e) {
            logger.error("解析失败",e);
           return WebUtil.errorResp(ErrCdEnum.C00990008.getCode(),ErrCdEnum.C00990008.getMsg(),null);
        }finally {
            if(outputStream != null){
                try {
                    outputStream.flush();
                    outputStream.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
        return WebUtil.successResp(null,"下载成功");
    }


    @RequestMapping("export")
    @ResponseBody
    public String export(HttpServletResponse response,@RequestParam Map<String,String> params){
        ServletOutputStream outputStream = null;
        try{
            response.setContentType("application/binary;charset=UTF-8");
            response.setHeader("Content-Disposition", "attachment;fileName=" + URLEncoder.encode("【dataDemo】" + ".xlsx", "UTF-8"));
            outputStream = response.getOutputStream();
            if("10".equals(params.get("toolCd"))){
                SXSSFWorkbook sxssfWorkbook = exportByPOI();
                sxssfWorkbook.write(outputStream);
            }
            if("20".equals(params.get("toolCd"))){
                String fileName = exportByEasy();
                File file = new File(fileName);
                IOUtils.copy(new FileInputStream(file),outputStream);
            }
        }catch (Exception e){
            logger.error("解析失败",e);
            return WebUtil.errorResp(ErrCdEnum.C00990008.getCode(),ErrCdEnum.C00990008.getMsg(),null);
        }finally {
            if(outputStream != null){
                try {
                    outputStream.flush();
                    outputStream.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
        return WebUtil.successResp(null,"下载成功");
    }

    private String analysisByPOI(MultipartFile multipartFile) throws IOException {
        InputStream inputStream = multipartFile.getInputStream();
        String fileName = multipartFile.getOriginalFilename();
        List<List<Object>> list = ExcelUtil.readExcel(inputStream,fileName.substring(fileName.lastIndexOf(".")+1),0);
        return JSON.toJSONString(list);
    }

    private String analysisByEasy(MultipartFile multipartFile) throws IOException {
        List<DemoData> retList = new ArrayList<>();
        InputStream inputStream = multipartFile.getInputStream();
        DemoDataListener demoDataListener = new DemoDataListener(retList);
        EasyExcel.read(inputStream,DemoData.class,demoDataListener).sheet().doRead();
        return JSON.toJSONString(retList);
    }

    private SXSSFWorkbook exportByPOI(){
        SXSSFWorkbook sxssfWorkbook = new SXSSFWorkbook();
        ExcelUtil.exportExcel(sxssfWorkbook);
        return sxssfWorkbook;
    }

    private String exportByEasy(){
        String fileName = TestFileUtil.getPath() + "file/dataDemo" + System.currentTimeMillis() + ".xlsx";
        EasyExcel.write(fileName, DemoData.class).sheet("模板").doWrite(WebUtil.data());
        return fileName;
    }
}
