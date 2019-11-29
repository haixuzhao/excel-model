package com.lanren.excelmodel.poi;

import com.alibaba.fastjson.JSON;
import com.lanren.excelmodel.util.ErrCdEnum;
import com.lanren.excelmodel.util.ExcelUtil;
import com.lanren.excelmodel.util.WebUtil;
import org.apache.tomcat.util.http.fileupload.disk.DiskFileItem;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.multipart.commons.CommonsMultipartFile;

import java.io.IOException;
import java.io.InputStream;
import java.util.List;
import java.util.Map;

/**
 * @ClassName POIExcelController
 * @Description: TODO
 * @Author zhx
 * @Date 2019/11/28
 * @Version V1.0
 **/
@Controller
@RequestMapping("excel/*")
public class POIExcelController {
    //定义一个全局的记录器，通过LoggerFactory获取
    private final static Logger logger = LoggerFactory.getLogger(POIExcelController.class);

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

    public String analysisByPOI(MultipartFile multipartFile) throws IOException {
        InputStream inputStream = multipartFile.getInputStream();
        String fileName = multipartFile.getOriginalFilename();
        List<List<Object>> list = ExcelUtil.readExcel(inputStream,fileName.substring(fileName.lastIndexOf(".")+1),0);
        return JSON.toJSONString(list);
    }

    public String analysisByEasy(MultipartFile multipartFile){
        throw new RuntimeException("异常测试");
    }

}
