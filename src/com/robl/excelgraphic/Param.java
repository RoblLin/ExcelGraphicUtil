package com.robl.excelgraphic;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;
import java.util.Map;

import javax.servlet.http.HttpServletResponse;

/**
 * @param dataSheetNum      操作sheet编号，默认0
 * @param dataMap           变量赋值
 * @param dataList          循环列表-目前循环里最多支持一个公式
 * @param expExcelName      生成Excel文件名称
 * @param templateExcepPath 模板Excel路径
 * @param editable          导出Excel是否可编辑
 * @param defaultPwd        设置不可编辑时设置密码
 * @param hiddenColumns     设置隐藏列
 * @param outputStream      设置输出流-必需
 * @param response          Web响应
 * @author Robl
 * @throws Exception
 */

public class Param {

    private int dataSheetNum;
    private Map<String, Object> dataMap;
    private List<Object> dataList;
    private String expExcelName;
    private InputStream templateInputStream;
    private boolean editable;
    private String defaultPwd;
    private int[] hiddenColumns;
    private OutputStream outputStream;
    private HttpServletResponse response;

    public Param() {
        dataSheetNum = 0;
        editable = true;
    }

    public int getDataSheetNum() {
        return dataSheetNum;
    }

    public void setDataSheetNum(int sheetNum) {
        this.dataSheetNum = sheetNum;
    }

    public String getDefaultPwd() {
        return defaultPwd;
    }

    public void setDefaultPwd(String defaultPwd) {
        this.defaultPwd = defaultPwd;
    }

    public boolean isEditable() {
        return editable;
    }

    public void setEditable(boolean editable) {
        this.editable = editable;
    }

    public String getExpExcelName() {
        return expExcelName;
    }

    public void setExpExcelName(String expExcelName) {
        this.expExcelName = expExcelName;
    }

    public InputStream getTemplateInputStream() {
        return templateInputStream;
    }

    public void setTemplateInputStream(InputStream templateInputStream) {
        this.templateInputStream = templateInputStream;
    }

    public int[] getHiddenColumns() {
        return hiddenColumns;
    }

    public void setHiddenColumns(int[] hiddenColumns) {
        this.hiddenColumns = hiddenColumns;
    }

    public OutputStream getOutputStream() {
        return outputStream;
    }

    public void setOutputStream(OutputStream outputStream) {
        this.outputStream = outputStream;
    }

    public HttpServletResponse getResponse() {
        return response;
    }

    public void setResponse(HttpServletResponse response) {
        this.response = response;
        try {
            if (expExcelName == null) {
                expExcelName = "导出Excel";
            }
            this.outputStream = response.getOutputStream();
            DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd HHmmss");
            String name = expExcelName.substring(0, expExcelName.lastIndexOf("."));
            name = new String(name.getBytes("UTF-8"), "ISO-8859-1");
            String suffix = expExcelName.substring(expExcelName.lastIndexOf("."));
            response.setHeader("Content-disposition",
                    "attachment;filename=" + name + dateFormat.format(new Date()) + suffix);
            response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8");
        } catch (IOException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
    }

    public Map<String, Object> getHeadMap() {
        return dataMap;
    }

    public void setHeadMap(Map<String, Object> headMap) {
        this.dataMap = headMap;
    }

    public List<Object> getDataList() {
        return dataList;
    }

    public void setDataList(List<Object> dataList) {
        this.dataList = dataList;
    }
}
