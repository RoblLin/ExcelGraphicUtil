package com.robl.excelgraphic;


import com.alibaba.fastjson.JSONObject;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.servlet.http.HttpServletResponse;
import java.io.InputStream;
import java.io.OutputStream;
import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.Map.Entry;

/**
 * Excel图表生成工具
 * 目前已不支持Excel03
 *
 * @author Robl
 */
public class OfficeUtils {

    private static final String START_LOOP_FLAG = "#LoopFlag#";
    private static final short RANGE_OF_DOUBLE_OF_XLS03_START = 170;
    private static final short RANGE_OF_DOUBLE_OF_XLS03_END = 200;
    private static final short RANGE_OF_NUMBER_OF_XLS03_1 = 177;
    private static final short RANGE_OF_NUMBER_OF_XLS03_2 = 178;
    private static final short RANGE_OF_NUMBER_OF_XLS03_3 = 182;

    /**
     * 按Excel模板，将List、Map导出Excel
     * <p>
     * #可支持循环变量及公式自动计算
     * <p>
     * #Tip：单元格内容为[日期：DATE]，变量DATE无法被赋值，需为[DATE]
     *
     * @param param
     * @throws Exception
     */
    public static void listToExcel(Param param) throws Exception {
        long t0 = System.currentTimeMillis();
        InputStream templateInputStream = param.getTemplateInputStream();
        if (templateInputStream == null) {
            System.out.println("模板Excel输入流不能为空");
            throw new Exception("模板Excel输入流不能为空! getTemplateInputStream()=null ");
        }

// 创建Excel文档
        Workbook workbook = null;
        boolean excel07Flag = false;
        try {
            workbook = paresExcel(templateInputStream);
        } catch (Exception e) {
            System.out.println("打开模板流失败!");
            throw new Exception("打开模板Excel流失败!templatePath=");
        }

        if (workbook instanceof XSSFWorkbook) {
            excel07Flag = true;
        }

        String fileName = param.getExpExcelName();
        if (fileName == null || "".equals(fileName)) {
            fileName = "导出报表.xlsx";
        } else {
            fileName = parseExcelName(fileName);
            if (excel07Flag && fileName.endsWith("xls")) {
                fileName += "x";// 如果是07版，未传入生成文件后缀，需要补上
            }
        }

//        long t2 = System.currentTimeMillis();
//        System.out.println("获取到workbook，耗时" + (t2 - t0) + "ms");

// 获取工作簿sheet
        int num = param.getDataSheetNum();
        Sheet sheet = workbook.getSheetAt(num);// 默认取第一个子表---此功能后续完善
// 获取总数
        int countRowNum = sheet.getLastRowNum();
        Row startLoopRow = null;

// 查找循环变量所在行
        for (int i = 0; i <= countRowNum; i++) {
            Row row = sheet.getRow(i);
            if (row == null) {
                continue;
            }
            for (int j = 0; j <= row.getLastCellNum(); j++) {// 循环每一列
                if (row.getCell(j) == null) {// 跳过单元格为空的情况
                    continue;
                }
                String cellValue = row.getCell(j).toString().trim();
                if (START_LOOP_FLAG.equals(cellValue)) {// 如果查询到startLoop，暂时保存一下
                    startLoopRow = row;
                    break;
                }
            }
        }

// 赋值非循环部分
        Map<String, Object> dateMap = param.getHeadMap();
        if (dateMap != null) {
            for (Entry<String, Object> entry : dateMap.entrySet()) {
                for (int i = 0; i <= countRowNum; i++) {
                    Row row = sheet.getRow(i);
                    if (row == null) {
                        continue;
                    }
                    for (int j = 0; j <= row.getLastCellNum(); j++) {// 循环每一列
                        if (row.getCell(j) == null) {// 跳过单元格为空的情况
                            continue;
                        }
                        String cellValue = row.getCell(j).toString().trim();
                        String entryKey = entry.getKey();
                        if (entryKey.startsWith("=")) {// 公式
                            cellValue = "=" + cellValue;// 从Excel获取时没有=
                        }
                        if (entryKey.equals(cellValue)) {// 赋值
                            CellStyle cellStyle = row.getCell(j).getCellStyle();
                            Cell newCell = row.getCell(j);
                            Object objValue = entry.getValue();

                            if (objValue instanceof Integer) {
                                newCell.setCellValue((Integer) objValue);
                            } else if (objValue instanceof Double) {
                                newCell.setCellValue((Double) objValue);
                            } else if (objValue instanceof Date) {
                                newCell.setCellValue(new SimpleDateFormat("yyyy-MM-dd").format((Date) objValue));
                            } else {

// 增加数值特殊化处理
                                Short cellType = cellStyle.getDataFormat();
                                if (cellType > RANGE_OF_DOUBLE_OF_XLS03_START
                                        && cellType < RANGE_OF_DOUBLE_OF_XLS03_END) {
// 数值格式
                                    if (cellType == RANGE_OF_NUMBER_OF_XLS03_1 || cellType == RANGE_OF_NUMBER_OF_XLS03_2
                                            || cellType == RANGE_OF_NUMBER_OF_XLS03_3) {
// 单元格格式设置为整数
                                        try {
                                            newCell.setCellValue(Integer.parseInt(newCell.getStringCellValue()));
                                        } catch (Exception e) {
                                            newCell.setCellValue(objValue + "");
                                        }
                                    } else {// 小数类型
                                        try {
                                            newCell.setCellValue(Double.parseDouble(newCell.getStringCellValue()));
                                        } catch (Exception e) {
                                            newCell.setCellValue(objValue + "");
                                        }
                                    }
                                } else {
                                    newCell.setCellValue(objValue + "");
                                }
                            }
                        }
                    }
                }
            }
        }

//        long t3 = System.currentTimeMillis();
//        System.out.println("赋值完所有headMap，耗时" + (t3 - t0) + "ms");

        int startLoopRowNum = -1;
        List<Object> dataList = param.getDataList();
        if (dataList != null && startLoopRow != null) {
            startLoopRowNum = startLoopRow.getRowNum();
            int listSize = 0;
// 赋值循环变量部分
            listSize = dataList.size();
            int k = startLoopRowNum;
            int startLoopColumNum = startLoopRow.getLastCellNum();
//            if (listSize != 0 && k < countRowNum) {
//                sheet.shiftRows(startLoopRowNum + 1, sheet.getLastRowNum(), listSize);// 先将循环下方内容向后平移listSize的行数，便于后边插入表格数据
//            }

            for (int i = 0; i < listSize; i++) {
                Row newRow = sheet.createRow(++k);// 新建一行
                newRow.setHeight(startLoopRow.getHeight());
                for (int j = 0; j < startLoopColumNum; j++) {
                    if (startLoopRow.getCell(j) == null) {// 跳过单元格为空的情况
                        continue;
                    }
                    CellStyle cellStyle = startLoopRow.getCell(j).getCellStyle();

                    String cellValue = startLoopRow.getCell(j).toString().trim();
                    Cell newCell = newRow.createCell(j);
                    newCell.setCellStyle(cellStyle);

                    JSONObject jsonObject = JSONObject.parseObject(JSONObject.toJSONString(dataList.get(i)));
                    Object obj = jsonObject.get(cellValue);

                    if (obj != null) {
                        if (cellValue.equals("#FORMAT")) {// 循环变量里的公式
                            String[] arr = (String[]) obj;
                            newCell.setCellFormula(arr[0].replaceFirst("=", ""));
                            newCell.setCellValue(arr[1]);
                        } else if (obj instanceof BigDecimal) {
                            newCell.setCellValue(((BigDecimal) obj).doubleValue());
                        } else if (obj instanceof Double) {
                            newCell.setCellValue((Double) obj);
                        } else if (obj instanceof Integer) {
                            newCell.setCellValue((Integer) obj);
                        } else if (obj instanceof Long) {//大多数情况，long类型为时间戳
                            SimpleDateFormat sft = new SimpleDateFormat("yyyy-MM-dd");
                            try {
                                newCell.setCellValue(sft.format(new Date((long) obj)));
                            } catch (Exception e) {
                                newCell.setCellValue(obj + "");
                            }
                        } else if (obj instanceof Date) {
                            SimpleDateFormat sft = new SimpleDateFormat("yyyy-MM-dd");
                            newCell.setCellValue(sft.format((Date) obj));
                        } else if (obj instanceof java.sql.Date) {
                            SimpleDateFormat sft = new SimpleDateFormat("yyyy-MM-dd");
                            newCell.setCellValue(sft.format((java.sql.Date) obj));
                        } else {
// 增加数值特殊化处理
                            Short cellType = cellStyle.getDataFormat();
                            if (cellType > RANGE_OF_DOUBLE_OF_XLS03_START && cellType < RANGE_OF_DOUBLE_OF_XLS03_END) {
// 数值格式
                                if (cellType == RANGE_OF_NUMBER_OF_XLS03_1 || cellType == RANGE_OF_NUMBER_OF_XLS03_2
                                        || cellType == RANGE_OF_NUMBER_OF_XLS03_3) {
// 单元格格式设置为整数
                                    try {
                                        newCell.setCellValue(Integer.parseInt(newCell.getStringCellValue()));
                                    } catch (Exception e) {
                                        newCell.setCellValue(obj + "");
                                    }
                                } else {// 小数类型
                                    try {
                                        newCell.setCellValue(Double.parseDouble(newCell.getStringCellValue()));
                                    } catch (Exception e) {
                                        newCell.setCellValue(obj + "");
                                    }
                                }
                            } else {
                                newCell.setCellValue(obj + "");
                            }
                        }
                    }
                }
            }
        }

        long t4 = System.currentTimeMillis();
        System.out.println("赋值完所有数据，耗时" + (t4 - t0) + "ms");

        if (startLoopRow != null) {
//            sheet.removeRow(startLoopRow);
// 删除变量名所在行，并将下面行上移
            if (startLoopRowNum > -1 && startLoopRow.getRowNum() < sheet.getLastRowNum()) {
                sheet.shiftRows(startLoopRow.getRowNum() + 1, sheet.getLastRowNum(), -1);
            }
        }

//        long t5 = System.currentTimeMillis();
//        System.out.println("所有list上移完成，耗时" + (t5 - t0) + "ms");

        List<int[]> autoMergeArea = param.getAutoMergeAreas();
        if (dataList != null && autoMergeArea != null && autoMergeArea.size() > 0) {

            List<int[]> referenceAreas = new ArrayList<>();
            int[] templateArea = new int[]{};
            for (int i = 0; i < autoMergeArea.size(); i++) {
                if (autoMergeArea.get(i).length == 1) {
                    templateArea = autoMergeArea.get(i);
                } else {
                    referenceAreas.add(autoMergeArea.get(i));
                }
            }

            //先合并参考行
            int lastIndex = -1;
            String template = "";
            for (int j = startLoopRowNum; j < dataList.size() + startLoopRowNum; j++) {//遍历每一行
                int[] arr = templateArea;
                String cellValue = "";
                try {
                    cellValue = sheet.getRow(j).getCell(arr[0]).getStringCellValue();
                } catch (Exception e) {
                    cellValue = sheet.getRow(j).getCell(arr[0]).getNumericCellValue() + "";
                }//暂时只支持到String和int两种类型

                if (cellValue == null || "".equals(cellValue)) {
                    lastIndex = j + 1;
                    template = "";
                    continue;
                }
                if (!cellValue.equals(template)) {
                    //把上面的合并
                    if (lastIndex > -1 && j - 1 > lastIndex) {
                        sheet.addMergedRegion(new CellRangeAddress(lastIndex, j - 1, arr[0], arr[0]));
                        CellStyle cellStyle = sheet.getRow(lastIndex).getCell(arr[0]).getCellStyle();
                        cellStyle.setAlignment(HorizontalAlignment.CENTER);
                        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
                        sheet.getRow(lastIndex).getCell(arr[0]).setCellStyle(cellStyle);
                    }
                    template = cellValue;
                    lastIndex = j;
                }
                if (j == dataList.size() + startLoopRowNum - 1 && lastIndex != j) {
                    sheet.addMergedRegion(new CellRangeAddress(lastIndex, j, arr[0], arr[0]));
                    CellStyle cellStyle = sheet.getRow(lastIndex).getCell(arr[0]).getCellStyle();
                    cellStyle.setAlignment(HorizontalAlignment.CENTER);
                    cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
                    sheet.getRow(lastIndex).getCell(arr[0]).setCellStyle(cellStyle);
                }
            }

            Map<Integer, Map<String, Object>> map = allMergedReferenceColumnRegion(sheet);
            for (int i = 0; i < referenceAreas.size(); i++
            ) {
                for (int j = startLoopRowNum; j < dataList.size() + startLoopRowNum; j++) {//遍历每一行
                    int[] arr = referenceAreas.get(i);
                    if (map.get(j) != null) {
                        sheet.addMergedRegion(new CellRangeAddress((Integer) map.get(j).get("firstRow"), (Integer) map.get(j).get("lastRow"), arr[0], arr[0]));
                        CellStyle cellStyle = sheet.getRow((Integer) map.get(j).get("firstRow")).getCell(arr[0]).getCellStyle();
                        cellStyle.setAlignment(HorizontalAlignment.CENTER);
                        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
                        sheet.getRow((Integer) map.get(j).get("firstRow")).getCell(arr[0]).setCellStyle(cellStyle);
                        j = (Integer) map.get(j).get("lastRow") + 1;
                    }
                }
            }

            long t6 = System.currentTimeMillis();
            System.out.println("合并完所有单元格，耗时" + (t6 - t0) + "ms");
        }
// 隐藏列
        Integer[] arr = param.getHiddenColumns();
        if (arr != null) {
            for (int i = 0; i < arr.length; i++) {
                sheet.setColumnHidden(arr[i], true);
            }
        }

// 设置是否可编辑
        boolean isEditable = param.isEditable();
        if (!isEditable) {
            String pwd = param.getDefaultPwd();
            if (pwd == null) {
                pwd = "";
            }
            sheet.protectSheet(pwd);// 此后可以设定密码
        }

// 公式自动计算
        sheet.setForceFormulaRecalculation(true);
// 创建文件输出流，准备输出电子表格
        OutputStream out;
        HttpServletResponse response = param.getResponse();
        if (response != null) {
            response.setContentType("application/vnd.ms-excel");
            response.addHeader("Content-Disposition", "attachment;filename=" + fileName);
            workbook.write(response.getOutputStream());
        } else {
            out = param.getOutputStream();
            workbook.write(out);
            out.flush();
            out.close();
            templateInputStream.close();
        }
        long t1 = System.currentTimeMillis();
        System.out.println("Excel导出成功！耗时：" + (t1 - t0) + "ms");


    }

    private static String parseExcelName(String fileName) {
        boolean flag = fileName.toLowerCase().endsWith(".xls") || fileName.toLowerCase().endsWith(".xlsx");
// 所有输入了后缀的情况
        if (flag) {// 将后缀统一转换为.xls--优化为可以支持07版
            int index = fileName.lastIndexOf(".");
            return fileName.substring(0, index) + fileName.substring(index).toLowerCase();
        } else {
            return fileName + ".xls";
        }
    }

    private static Workbook paresExcel(InputStream templateInputStream) throws Exception {
        return new XSSFWorkbook(templateInputStream);
    }

    private static Map<Integer, Map<String, Object>> allMergedReferenceColumnRegion(Sheet sheet) {
        int sheetMergeCount = sheet.getNumMergedRegions();
        Map<Integer, Map<String, Object>> mapTotal = new HashMap<>();
        for (int i = 0; i < sheetMergeCount; i++) {
            CellRangeAddress range = sheet.getMergedRegion(i);
            int firstColumn = range.getFirstColumn();
            int lastColumn = range.getLastColumn();
            int firstRow = range.getFirstRow();
            int lastRow = range.getLastRow();
            for (int j = firstRow; j <= lastRow; j++) {
                Map<String, Object> map = new HashMap<>();
                map.put("isMergedRegion", true);
                map.put("firstRow", firstRow);
                map.put("lastRow", lastRow);
                map.put("firstColumn", firstColumn);
                map.put("lastColumn", lastColumn);
                mapTotal.put(j, map);
            }
        }
        return mapTotal;
    }
}