package com.tiangang.poi.excel.template;

import lombok.Data;

import java.util.Collections;
import java.util.List;
import java.util.Map;

@Data
public class DynamicSource {
    private String id;
    private List<Map<String, String>> dataList;

    public static List<DynamicSource> createList(String id, List<Map<String, String>> dataList) {
        DynamicSource dynamicSource = new DynamicSource();
        dynamicSource.id = id;
        dynamicSource.dataList = dataList;
        return Collections.singletonList(dynamicSource);
    }


}
