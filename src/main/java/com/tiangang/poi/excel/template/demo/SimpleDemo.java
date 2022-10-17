package com.tiangang.poi.excel.template.demo;

import com.tiangang.poi.excel.template.ExcelTemplateUtil;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.InputStream;
import java.util.HashMap;
import java.util.Map;

public class SimpleDemo {
    public static void main(String[] args) throws Exception {
        Map<String, String> staticSource = new HashMap<>();
        staticSource.put("title", "poi-excel-template");
        // 1.从resources下加载模板并替换
        InputStream resourceAsStream = SimpleDemo.class.getClassLoader().getResourceAsStream("simple-template.xlsx");
        Workbook workbook = ExcelTemplateUtil.buildByTemplate(resourceAsStream, staticSource, null);
        // 2.保存到本地
        ExcelTemplateUtil.save(workbook, "D:\\simple-poi-excel-template.xlsx");
    }
}
