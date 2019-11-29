package com.lanren.excelmodel.util;

import java.awt.Color;
import java.awt.geom.IllegalPathStateException;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.UUID;
import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.util.StringUtils;
import com.alibaba.fastjson.JSONObject;

/**
 * @description 数据导出工具类
 */
public class ExcelUtil {
    /**
     * 日期格式化
     */
    private static final SimpleDateFormat DATE_FORMAT = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
    /**
     * 数字格式化
     */
    private static final DecimalFormat DECIMAL_FORMAT = new DecimalFormat("0.00");
    /**
     * 默认开始读取的行号
     */
    private static final int DEFAULT_START_ROW_NUM = 0;
    /**
     * excel2003扩展名
     */
    private static final String EXT_2003 = "xls";
    /**
     * excel2007扩展名
     */
    private static final String EXT_2007 = "xlsx";

	/**
	 * 每个sheet的行数
	 */
	private static final Integer SHEET_SIZE = 10000;

	public static enum Style {
		DEF, BLUE
	}

	/**
	 * @description 页面导出excel
	 * @author wangyongchao1
	 * @param titleList 标题列表<br/>
	 *                  key表示导出数据对应的字段名称，value表示标题显示的内容
	 * @param dataList  要导出的数据
	 * @param response
	 * @param fileName
	 * @throws Exception
	 */
	public static String exportExcel(List<Map<String, String>> titleList, List<Map<String, Object>> dataList, HttpServletResponse response, String fileName, Style style) {

		if (dataList == null || dataList.size() == 0) {
			JSONObject data = new JSONObject();
			data.put("success", false);
			data.put("message", "没有数据可以导出!");
			return data.toString();
		}
		try {
			// 第一步，创建一个workbook，对应一个Excel文件
			SXSSFWorkbook workbook = new SXSSFWorkbook(100);
			XSSFCellStyle titleStyle = getTitleStyle(workbook, style);
			// 创建内容数值样式
			XSSFCellStyle contentIntStyle = getContentIntStyle(workbook);
			// 创建内容字符串文本样式
			XSSFCellStyle contentStringStyle = getContentStringStyle(workbook);
			int sheetNumber = dataList.size() % SHEET_SIZE == 0 ? dataList.size() / SHEET_SIZE : dataList.size() / SHEET_SIZE + 1;

			for (int s = 1; s <= sheetNumber; s++) {
				// 第二步，在webbook中添加一个sheet,对应Excel文件中的sheet
				Sheet sheet = workbook.createSheet("sheet" + s);
				// 第三步，在sheet中添加表头第0行
				Row row = sheet.createRow(0);
//				// 第四步，创建单元格，并设置值表头 设置表头居中

				Cell cell = null;
				List<String> fieldList = new ArrayList<>();

				for (int i = 0; i < titleList.size(); i++) {
					Map<String, String> titleMap = titleList.get(i);
					for (Entry<String, String> entry : titleMap.entrySet()) {
						String key = entry.getKey();
						fieldList.add(key);
						String title = entry.getValue();
						cell = row.createCell(i);
						cell.setCellValue(title);
						cell.setCellStyle(titleStyle);
					}
				}

//				 第五步，写入数据
				Integer startIndex = (s - 1) * SHEET_SIZE;
				Integer endIndex = startIndex + SHEET_SIZE >= dataList.size() ? dataList.size() : startIndex + SHEET_SIZE;

				for (int i = startIndex, rowNumber = 0; i < endIndex; i++, rowNumber++) {
					row = sheet.createRow(rowNumber + 1);
					Map<String, Object> dataMap = dataList.get(i);
					for (int j = 0; j < fieldList.size(); j++) {
						Object content = dataMap.get(fieldList.get(j));
//						第六步，创建单元格，并设置值
						cell = row.createCell(j);
						if (checkNumber(content.toString())) {
							cell.setCellStyle(contentIntStyle);
						} else {
							cell.setCellStyle(contentStringStyle);
						}
						cell.setCellValue(content.toString());
					}
				}

			}
			// 第七步，将文件输出到客户端浏览器
			response.setContentType("application/binary;charset=UTF-8");
			if (StringUtils.isEmpty(fileName)) {
				fileName = UUID.randomUUID().toString().replace("-", "");
			}
			response.setHeader("Content-Disposition", "attachment;fileName=" + new String(fileName.getBytes("gbk"), "ISO-8859-1") + ".xlsx");
			ServletOutputStream out = response.getOutputStream();
			workbook.write(out);
			out.flush();
			out.close();

			JSONObject data = new JSONObject();
			data.put("success", true);
			data.put("message", "导出文件成功!");
			return data.toString();
		} catch (Exception e) {
			e.printStackTrace();
			JSONObject data = new JSONObject();
			data.put("success", false);
			data.put("message", "导出文件出现异常!");
			return data.toString();

		}
	}

	private static XSSFCellStyle getTitleStyle(SXSSFWorkbook workbook, Style style) {
//		创建标题样式
		XSSFCellStyle titleStyle = workbook.getXSSFWorkbook().createCellStyle();
		titleStyle.setAlignment(HorizontalAlignment.LEFT); 
		titleStyle.setVerticalAlignment(VerticalAlignment.CENTER);
		titleStyle.setBorderBottom(BorderStyle.THIN);
		titleStyle.setBorderLeft(BorderStyle.THIN);
		titleStyle.setBorderTop(BorderStyle.THIN);
		titleStyle.setBorderRight(BorderStyle.THIN);
		XSSFFont font = workbook.getXSSFWorkbook().createFont();
		font.setFontHeightInPoints((short) 12);
		titleStyle.setFont(font);
		titleStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		switch (style) {
		case BLUE:
			// 设置背景色 蓝色
			titleStyle.setFillForegroundColor(new XSSFColor(new Color(0, 177, 241)));
			break;
		case DEF:
			// 设置背景色 灰色
			titleStyle.setFillForegroundColor(new XSSFColor(new Color(229, 223, 237)));
			break;
		default:
			// 设置默认背景色 灰色
			titleStyle.setFillForegroundColor(new XSSFColor(new Color(229, 223, 237)));
			break;
		}
		return titleStyle;
	}

	private static XSSFCellStyle getContentIntStyle(SXSSFWorkbook workbook) {
		XSSFCellStyle contentIntStyle = workbook.getXSSFWorkbook().createCellStyle();
		contentIntStyle.setAlignment(HorizontalAlignment.RIGHT);
		contentIntStyle.setVerticalAlignment(VerticalAlignment.CENTER);
		contentIntStyle.setBorderBottom(BorderStyle.THIN);
		contentIntStyle.setBorderLeft(BorderStyle.THIN);
		contentIntStyle.setBorderTop(BorderStyle.THIN);
		contentIntStyle.setBorderRight(BorderStyle.THIN);
		return contentIntStyle;
	}

	private static XSSFCellStyle getContentStringStyle(SXSSFWorkbook workbook) {
		XSSFCellStyle contentStringStyle = workbook.getXSSFWorkbook().createCellStyle();
		contentStringStyle.setAlignment(HorizontalAlignment.LEFT);
		contentStringStyle.setVerticalAlignment(VerticalAlignment.CENTER);
		contentStringStyle.setBorderBottom(BorderStyle.THIN);
		contentStringStyle.setBorderLeft(BorderStyle.THIN);
		contentStringStyle.setBorderTop(BorderStyle.THIN);
		contentStringStyle.setBorderRight(BorderStyle.THIN);
		return contentStringStyle;
	}

	/**
	 * 判断一段字符串是否为数字
	 * 
	 * @param str
	 * @return
	 */
	private static Boolean checkNumber(String str) {
		String regex = "-[0-9]+(.[0-9]+)?|[0-9]+(.[0-9]+)?";
		if (str == null || !str.matches(regex) || "".equals(str)) {
			return false;
		}
		return true;
	}
	/**
     * @description 读取Excel数据
     * @author wangyongchao1
     * @param in
     * @param extension   扩展名xsl或者xslx
     * @param startRowNum 开始读取的行号
     * @return excel行数据
     * @throws IOException
     */
    public static List<List<Object>> readExcel(InputStream in, String extension, Integer startRowNum)
            throws IOException {
        if (startRowNum == null || startRowNum < 0) {
            startRowNum = DEFAULT_START_ROW_NUM;
        }
        if (EXT_2003.equals(extension)) {
            return readExcel2003(in, startRowNum);
        } else if (EXT_2007.equals(extension)) {
            return readExcel2007(in, startRowNum);
        } else {
            throw new IllegalPathStateException("不支持的文件类型");
        }
    }

    private static List<List<Object>> readExcel2003(InputStream in, Integer startRowNum) throws IOException {
        List<List<Object>> rowList = new ArrayList<>();
        HSSFWorkbook workbook = new HSSFWorkbook(in);
        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            HSSFSheet sheet = workbook.getSheetAt(i);
            int lastRowNum = sheet.getLastRowNum();
            for (int ii = startRowNum; ii <= lastRowNum; ii++) {
                HSSFRow row = sheet.getRow(ii);
                int lastCellNum = row.getLastCellNum();
                List<Object> cellList = new ArrayList<>();
                rowList.add(cellList);
                for (int iii = 0; iii <= lastCellNum; iii++) {
                    HSSFCell cell = row.getCell(iii);
                    addCellValue(cell, cellList);
                }
            }
        }
        workbook.close();
        return rowList;
    }

    private static List<List<Object>> readExcel2007(InputStream in, Integer startRowNum) throws IOException {
        List<List<Object>> rowList = new ArrayList<>();
        XSSFWorkbook workbook = new XSSFWorkbook(in);
        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            XSSFSheet sheet = workbook.getSheetAt(i);
            int lastRowNum = sheet.getLastRowNum();
            for (int ii = startRowNum; ii <= lastRowNum; ii++) {
                XSSFRow row = sheet.getRow(ii);
                int lastCellNum = row.getLastCellNum();
                List<Object> cellList = new ArrayList<>();
                rowList.add(cellList);
                for (int iii = 0; iii <= lastCellNum; iii++) {
                    XSSFCell cell = row.getCell(iii);
                    addCellValue(cell, cellList);
                }
            }
        }
        workbook.close();
        return rowList;
    }

    private static void addCellValue(Cell cell, List<Object> cellList) {
        if (cell == null) {
            cellList.add(null);
            return;
        }
        switch (cell.getCellTypeEnum()) {
        case BOOLEAN:
            cellList.add(cell.getBooleanCellValue());
            break;
        case BLANK:
            cellList.add(null);
            break;
        case NUMERIC:
            if (DateUtil.isCellDateFormatted(cell)) {
                cellList.add(DATE_FORMAT.format(cell.getDateCellValue()));
            } else {
                cellList.add(DECIMAL_FORMAT.format(cell.getNumericCellValue()));
            }
            break;
        default:
            cellList.add(cell.getStringCellValue());
            break;
        }
    }

    public static void main(String[] args) {
        File file = new File("d://发票单20190118.xls");
        try {
            List<List<Object>> dataList = readExcel(new FileInputStream(file), "xls", 4);
            for (List<Object> list : dataList) {
                System.out.println(list);
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
