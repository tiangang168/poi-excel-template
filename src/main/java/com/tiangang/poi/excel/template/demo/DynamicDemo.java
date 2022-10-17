package com.tiangang.poi.excel.template.demo;

import com.tiangang.poi.excel.template.DynamicSource;
import com.tiangang.poi.excel.template.ExcelTemplateUtil;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class DynamicDemo {
    public static void main(String[] args) throws Exception {

        Map<String, String> staticSource = new HashMap<>();
        staticSource.put("title", "poi-excel-template");
        // 模拟10行
        int rows = 10;
        List<Map<String, String>> dataList = new ArrayList<>();
        for (int i = 1; i <= rows; i++) {
            // 一行
            Map<String, String> rowMap = new HashMap<>();
            rowMap.put("id", "" + i);
            rowMap.put("name", "name" + i);
            rowMap.put("price", "" + (i * 100));
            rowMap.put("unit", "unit" + i);
            rowMap.put("discount", "" + i);
            rowMap.put("sellingPrice", "" + (i * 100 - 10));
            dataList.add(rowMap);
        }
        // 可以创建多个id，这里只创建1个示例
        List<DynamicSource> dynamicSourceList = DynamicSource.createList("p", dataList);
        // 1.从resources下加载模板并替换
        InputStream resourceAsStream = DynamicDemo.class.getClassLoader().getResourceAsStream("dynamic-template.xlsx");
        Workbook workbook = ExcelTemplateUtil.buildByTemplate(resourceAsStream, staticSource, dynamicSourceList);
        // 2.保存到本地
        ExcelTemplateUtil.save(workbook, "D:\\dynamic-poi-excel-template.xlsx");
    }
}
