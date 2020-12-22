package com.robl.excelgraphic;

import javax.servlet.http.HttpServletResponse;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;
import java.util.Map;

/**
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
    private Integer[] hiddenColumns;
    private OutputStream outputStream;
    private HttpServletResponse response;
    private List<int[]> autoMergeAreas;

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

    public Integer[] getHiddenColumns() {
        return hiddenColumns;
    }

    public void setHiddenColumns(Integer[] hiddenColumns) {
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

    public List<int[]> getAutoMergeAreas() {
        return autoMergeAreas;
    }

    public void setAutoMergeAreas(List<int[]> autoMergeAreas) {
        this.autoMergeAreas = autoMergeAreas;
    }
}
