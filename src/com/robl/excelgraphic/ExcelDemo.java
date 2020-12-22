package com.robl.excelgraphic;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.*;

public class ExcelDemo {

    public static void main(String[] args) throws Exception {
        System.out.println("资源文件存放于此：" + System.getProperty("user.dir"));
        String templateXls = System.getProperty("user.dir") + "\\template.xlsx";
        String destXls = System.getProperty("user.dir") + "\\statistics.xlsx";
        Param param = new Param();
        param.setTemplateInputStream(new FileInputStream(new File(templateXls)));
        param.setOutputStream(new FileOutputStream(new File(destXls)));

        Map<String, Object> dataMap = new HashMap<>();
        dataMap.put("date", new Date());
        dataMap.put("operator", "萝卜很大");
        param.setHeadMap(dataMap);

        List<Map<String, Object>> destData = new ArrayList<>();
        for (int i = 0; i < 12; i++) {
            HashMap<String, Object> hashMap = new HashMap<>();
            hashMap.put("month", (i + 1) + "月");
            hashMap.put("p0", 6.4 * Math.random() * 10);
            hashMap.put("n0", (int) (4 * Math.random() * 20));
            hashMap.put("p1", 18.4 * Math.random() * 10);
            hashMap.put("n1", (int) (34 * Math.random() * 20));
            hashMap.put("p2", 4.4 * Math.random() * 10);
            hashMap.put("n2", (int) (14 * Math.random() * 20));
            destData.add(hashMap);
        }
        param.setDataList(Collections.unmodifiableList(destData));
        OfficeUtils.listToExcel(param);
    }
}