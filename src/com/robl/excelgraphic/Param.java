package com.robl.excelgraphic;

import java.io.IOException;
import java.io.OutputStream;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;
import java.util.Map;

import javax.servlet.http.HttpServletResponse;

/**
 * @param dataSheetNum
 *            ����sheet��ţ�Ĭ��0
 * @param dataMap
 *            ������ֵ
 * @param dataList
 *            ѭ���б�-Ŀǰѭ�������֧��һ����ʽ
 * @param expExcelName
 *            ����Excel�ļ�����
 * @param templateExcepPath
 *            ģ��Excel·��
 * @param editable
 *            ����Excel�Ƿ�ɱ༭
 * @param defaultPwd
 *            ���ò��ɱ༭ʱ��������
 * @param hiddenColumns
 *            ����������
 * @param outputStream
 *            ���������-����
 * @param response
 *            Web��Ӧ
 * @throws Exception
 * 
 * @author Robl
 */

public class Param {

	private int dataSheetNum;
	private Map<String, Object> dataMap;
	private List<Map<String, Object>> dataList;
	private String expExcelName;
	private String templateExcepPath;
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

	public String getTemplateExcepPath() {
		return templateExcepPath;
	}

	public void setTemplateExcepPath(String templateExcepPath) {
		this.templateExcepPath = templateExcepPath;
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
				expExcelName = "����Excel";
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

	public List<Map<String, Object>> getDataList() {
		return dataList;
	}

	public void setDataList(List<Map<String, Object>> dataList) {
		this.dataList = dataList;
	}
}
