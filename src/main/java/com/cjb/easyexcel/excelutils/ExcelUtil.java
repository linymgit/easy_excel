package com.cjb.easyexcel.excelutils;

import com.github.crab2died.ExcelUtils;
import com.github.crab2died.exceptions.Excel4JException;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.net.URLEncoder;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.function.Function;

public class ExcelUtil {
    /**
     * 下载excel
     * @param request
     * @param response
     * @param fileName
     * @param header
     * @param data
     * @throws IOException
     */
    public static void exportExcel(HttpServletRequest request, HttpServletResponse response,
                                   String fileName, List<String> header, List<List<String>> data) throws IOException {
        response.setHeader("Content-Disposition", "attachment; filename=" + URLEncoder.encode(fileName, "UTF-8"));
        response.setContentType(request.getServletContext().getMimeType("xx.xlsx"));
        response.setContentType("application/octet-stream");
        ExcelUtils.getInstance().exportObjects2Excel(data, header, response.getOutputStream());
        response.getOutputStream().flush();
    }

    public static void exportExcel(FileOutputStream output, List<String> header, List<List<String>> data) throws IOException {
    	ExcelUtils.getInstance().exportObjects2Excel(data, header, output);
    	output.flush();
    }

    /**
     * 基于注解的data 导出Excel表格
     */
    public static void exportExcel(HttpServletRequest request, HttpServletResponse response, String fileName, Class<?> clazz, List<?> data)
            throws IOException, Excel4JException {
        response.setHeader("Content-Disposition", "attachment; filename=" + URLEncoder.encode(fileName, "UTF-8"));
        response.setContentType(request.getServletContext().getMimeType("xx.xlsx"));
        response.setContentType("application/octet-stream");
        ExcelUtils.getInstance().exportObjects2Excel(data, clazz, true, response.getOutputStream());
        response.getOutputStream().flush();
    }
    public static void exportExcel(OutputStream output, Class<?> clazz, List<?> data)throws IOException, Excel4JException {
    	ExcelUtils.getInstance().exportObjects2Excel(data, clazz, true, output);
    }
    public static void changeCellColorExcel(InputStream originExcel,OutputStream targetExcel,Function<Row,IndexedColors> row2Color
    		,Function<Cell,IndexedColors> cell2Color)throws IOException, Excel4JException, EncryptedDocumentException, InvalidFormatException {
    	Map<IndexedColors, CellStyle> color2StyleCache=new HashMap<>();
    	try (Workbook workbook = WorkbookFactory.create(originExcel)) {
    		Sheet sheet = workbook.getSheetAt(0);
    		long maxLine = sheet.getLastRowNum();
    		for (int i = 0; i <= maxLine; i++) {
    			Row row = sheet.getRow(i);
    			if (null == row){
    				continue;
    			}
    			IndexedColors color = row2Color.apply(row);
    			CellStyle rowStyle=color2StyleCache.get(color);
    			if(rowStyle==null){
    				rowStyle = workbook.createCellStyle();
    				rowStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    				rowStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    				rowStyle.setFillForegroundColor(color.getIndex());
    				rowStyle.setFillBackgroundColor(color.getIndex());
    				color2StyleCache.put(color, rowStyle);
    			}
    			for (Cell cell : row) {
    				IndexedColors cellColor = cell2Color==null?null:cell2Color.apply(cell);
    				if(cellColor!=null){
    					CellStyle cellStyle = workbook.createCellStyle();
    					cellStyle.cloneStyleFrom(rowStyle);
    					cellStyle.setFillForegroundColor(cellColor.getIndex());
    					cellStyle.setFillBackgroundColor(cellColor.getIndex());
    					cell.setCellStyle(cellStyle);
    				}else{
    					cell.setCellStyle(rowStyle);
    				}
    			}
    		}
    		workbook.write(targetExcel);
    	}
    }
}
