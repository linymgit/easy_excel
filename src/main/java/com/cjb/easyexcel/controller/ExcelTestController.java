package com.cjb.easyexcel.controller;


import com.cjb.easyexcel.entity.EntityTest;
import com.cjb.easyexcel.excelutils.ExcelUtil;
import com.github.crab2died.ExcelUtils;
import com.github.crab2died.exceptions.Excel4JException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

/**
 * @author cjb
 * @version 0.0.1
 * 上传、导出Excel
 */
@RestController
public class ExcelTestController {

    @GetMapping("/hello")
    public String sayHello(){
        return "hello cjb!!!";
    }


    /**
     * 上传Excel
     */
    @PostMapping("/upload")
    public synchronized Object uploadExcelTest(@RequestParam("file") MultipartFile file) {
        if (file.isEmpty()) {
            return "老哥，上传的文件不能为空";
        }

        try (InputStream is = file.getInputStream()) {

            ExcelUtils utils = ExcelUtils.getInstance();
            int sheetIndex = 0;
            // 实体类 （注解）看下例子就会用
            List<EntityTest> testData = utils.readExcel2Objects(is, EntityTest.class, sheetIndex);
            return testData;

        } catch (IOException e) {
            e.printStackTrace();
        } catch (InvalidFormatException e) {
            e.printStackTrace();
        } catch (Excel4JException e) {
            e.printStackTrace();
        }

        return null;
    }


    /**
     * 下载Excel
     */
    @RequestMapping("/export")
    public synchronized String exportExcel(HttpServletRequest request, HttpServletResponse response) {
        try {
            List<EntityTest> data = new ArrayList<>();
            for (int i = 0; i < 100; i++) {
                EntityTest test =  new EntityTest();
                test.setField1("1field"+i);
                test.setField2("2field"+i);
                data.add(test);
            }
            ExcelUtil.exportExcel(request,response,"cjb导出的Excel.xlsx", EntityTest.class, data);
        } catch (Exception ex) {
            return "老哥，怎么导出出问题了！！！";
        }
        return "老哥，恭喜你 导出成功";
    }
}