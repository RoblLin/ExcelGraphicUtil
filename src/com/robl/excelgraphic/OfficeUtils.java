package com.robl.excelgraphic;

import java.io.FileInputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Excelͼ�����ɹ���
 * 
 * @author Robl
 *
 */
public class OfficeUtils {

	private static final String START_LOOP_FLAG = "#LoopFlag#";
	private static final short RANGE_OF_DOUBLE_OF_XLS03_START = 170;
	private static final short RANGE_OF_DOUBLE_OF_XLS03_END = 200;
	private static final short RANGE_OF_NUMBER_OF_XLS03_1 = 177;
	private static final short RANGE_OF_NUMBER_OF_XLS03_2 = 178;
	private static final short RANGE_OF_NUMBER_OF_XLS03_3 = 182;

	/**
	 * ��Excelģ�壬��List��Map����Excel
	 * 
	 * #��֧��ѭ����������ʽ�Զ�����
	 * 
	 * #Tip����Ԫ������Ϊ[���ڣ�DATE]������DATE�޷�����ֵ����Ϊ[DATE]
	 * 
	 * @param param
	 * @throws Exception
	 */
	public static void listToExcel(Param param) throws Exception {
		long t0 = System.currentTimeMillis();
		String templatePath = param.getTemplateExcepPath();
		if (templatePath == null || "".equals(templatePath)) {
			System.out.println("ģ��Excel����Ϊ��");
			throw new Exception("ģ��Excel·������Ϊ��! getTemplatePath()=null ");
		}

		// ����Excel�ĵ�
		Workbook workbook = null;
		boolean excel07Flag = false;
		try {
			workbook = paresExcel(templatePath);
		} catch (Exception e) {
			System.out.println("��ģ��ʧ��!");
			throw new Exception("��ģ��Excelʧ��!templatePath=" + templatePath);
		}

		if (workbook instanceof XSSFWorkbook) {
			excel07Flag = true;
		}

		String fileName = param.getExpExcelName();
		if (fileName == null || "".equals(fileName)) {
			fileName = "��������.xls";
		} else {
			fileName = parseExcelName(fileName);
			if (excel07Flag && fileName.endsWith("xls")) {
				fileName += "x";// �����07�棬δ���������ļ���׺����Ҫ����
			}
		}

		// ��ȡ������sheet
		int num = param.getDataSheetNum();
		Sheet sheet = workbook.getSheetAt(num);// Ĭ��ȡ��һ���ӱ�---�˹��ܺ�������
		// ��ȡ����
		int countRowNum = sheet.getLastRowNum();
		Row startLoopRow = null;

		// ����ѭ������������
		for (int i = 0; i <= countRowNum; i++) {
			Row row = sheet.getRow(i);
			if (row == null) {
				continue;
			}
			for (int j = 0; j <= row.getLastCellNum(); j++) {// ѭ��ÿһ��
				if (row.getCell(j) == null) {// ������Ԫ��Ϊ�յ����
					continue;
				}
				String cellValue = row.getCell(j).toString().trim();
				if (START_LOOP_FLAG.equals(cellValue)) {// �����ѯ��startLoop����ʱ����һ��
					startLoopRow = row;
					break;
				}
			}
		}

		// ��ֵ��ѭ������
		Map<String, Object> dateMap = param.getHeadMap();
		if (dateMap != null) {
			for (Entry<String, Object> entry : dateMap.entrySet())
				for (int i = 0; i <= countRowNum; i++) {
					Row row = sheet.getRow(i);
					if (row == null) {
						continue;
					}
					for (int j = 0; j <= row.getLastCellNum(); j++) {// ѭ��ÿһ��
						if (row.getCell(j) == null) {// ������Ԫ��Ϊ�յ����
							continue;
						}
						String cellValue = row.getCell(j).toString().trim();
						String entryKey = entry.getKey();
						if (entryKey.startsWith("=")) {// ��ʽ
							cellValue = "=" + cellValue;// ��Excel��ȡʱû��=
						}
						if (entryKey.equals(cellValue)) {// ��ֵ
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

								// ������ֵ���⻯����
								Short cellType = cellStyle.getDataFormat();
								if (cellType > RANGE_OF_DOUBLE_OF_XLS03_START
										&& cellType < RANGE_OF_DOUBLE_OF_XLS03_END) {
									// ��ֵ��ʽ
									if (cellType == RANGE_OF_NUMBER_OF_XLS03_1 || cellType == RANGE_OF_NUMBER_OF_XLS03_2
											|| cellType == RANGE_OF_NUMBER_OF_XLS03_3) {
										// ��Ԫ���ʽ����Ϊ����
										try {
											newCell.setCellValue(Integer.parseInt(newCell.getStringCellValue()));
										} catch (Exception e) {
											newCell.setCellValue(objValue + "");
										}
									} else {// С������
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

		int startLoopRowNum = -1;
		List<Map<String, Object>> dataList = param.getDataList();
		if (dataList != null && startLoopRow != null) {
			startLoopRowNum = startLoopRow.getRowNum();
			int listSize = 0;
			// ��ֵѭ����������
			listSize = dataList.size();
			int k = startLoopRowNum;
			int startLoopColumNum = startLoopRow.getLastCellNum();
			if (listSize != 0) {
				sheet.shiftRows(startLoopRowNum + 1, countRowNum, listSize);// �Ƚ�ѭ���·��������ƽ��listSize�����������ں�߲���������
			}

			for (int i = 0; i < listSize; i++) {
				Row newRow = sheet.createRow(++k);// �½�һ��
				newRow.setHeight(startLoopRow.getHeight());
				for (int j = 0; j < startLoopColumNum; j++) {
					if (startLoopRow.getCell(j) == null) {// ������Ԫ��Ϊ�յ����
						continue;
					}
					CellStyle cellStyle = startLoopRow.getCell(j).getCellStyle();

					String cellValue = startLoopRow.getCell(j).toString().trim();
					Cell newCell = newRow.createCell(j);
					newCell.setCellStyle(cellStyle);

					Object obj = dataList.get(i).get(cellValue);

					if (obj != null) {
						if (cellValue.equals("#FORMAT")) {// ѭ��������Ĺ�ʽ
							String[] arr = (String[]) obj;
							newCell.setCellFormula(arr[0].replaceFirst("=", ""));
							newCell.setCellValue(arr[1]);
						} else if (obj instanceof BigDecimal) {
							newCell.setCellValue(((BigDecimal) obj).doubleValue());
						} else if (obj instanceof Double) {
							newCell.setCellValue((Double) obj);
						} else if (obj instanceof Integer) {
							newCell.setCellValue((Integer) obj);
						} else if (obj instanceof Date) {
							SimpleDateFormat sft = new SimpleDateFormat("yyyy-MM-dd");
							newCell.setCellValue(sft.format((Date) obj));
						} else if (obj instanceof java.sql.Date) {
							SimpleDateFormat sft = new SimpleDateFormat("yyyy-MM-dd");
							newCell.setCellValue(sft.format((java.sql.Date) obj));
						} else {
							// ������ֵ���⻯����
							Short cellType = cellStyle.getDataFormat();
							if (cellType > RANGE_OF_DOUBLE_OF_XLS03_START && cellType < RANGE_OF_DOUBLE_OF_XLS03_END) {
								// ��ֵ��ʽ
								if (cellType == RANGE_OF_NUMBER_OF_XLS03_1 || cellType == RANGE_OF_NUMBER_OF_XLS03_2
										|| cellType == RANGE_OF_NUMBER_OF_XLS03_3) {
									// ��Ԫ���ʽ����Ϊ����
									try {
										newCell.setCellValue(Integer.parseInt(newCell.getStringCellValue()));
									} catch (Exception e) {
										newCell.setCellValue(obj + "");
									}
								} else {// С������
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
		if (startLoopRow != null) {
			sheet.removeRow(startLoopRow);
			// ɾ�������������У���������������
			sheet.shiftRows(startLoopRowNum + 1, sheet.getLastRowNum(), -1);
		}

		// ������
		int[] arr = param.getHiddenColumns();
		if (arr != null) {
			for (int i = 0; i < arr.length; i++) {
				sheet.setColumnHidden(arr[i], true);
			}
		}

		// �����Ƿ�ɱ༭
		boolean isEditable = param.isEditable();
		if (!isEditable) {
			String pwd = param.getDefaultPwd();
			if (pwd == null) {
				pwd = "";
			}
			sheet.protectSheet(pwd);// �˺�����趨����
		}

		// ��ʽ�Զ�����
		sheet.setForceFormulaRecalculation(true);
		// �����ļ��������׼��������ӱ��
		OutputStream out = param.getOutputStream();
		if (out == null) {
			throw new Exception("���������OutputStream����Ϊ�գ�");
		}
		workbook.write(out);
		out.flush();
		out.close();
		long t1 = System.currentTimeMillis();
		System.out.println("Excel�����ɹ�����ʱ��" + (t1 - t0) + "ms");
	}

	private static String parseExcelName(String fileName) {
		boolean flag = fileName.toLowerCase().endsWith(".xls") || fileName.toLowerCase().endsWith(".xlsx");
		// ���������˺�׺�����
		if (flag) {// ����׺ͳһת��Ϊ.xls--�Ż�Ϊ����֧��07��
			int index = fileName.lastIndexOf(".");
			return fileName.substring(0, index) + fileName.substring(index).toLowerCase();
		} else {
			return fileName + ".xls";
		}
	}

	private static Workbook paresExcel(String excelpath) throws Exception {
		InputStream in = new FileInputStream(excelpath);
		return excelpath.endsWith("xlsx") ? new XSSFWorkbook(in) : new HSSFWorkbook(new POIFSFileSystem(in));
	}
}