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
 * Excel图表生成工具
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
	 * 按Excel模板，将List、Map导出Excel
	 * 
	 * #可支持循环变量及公式自动计算
	 * 
	 * #Tip：单元格内容为[日期：DATE]，变量DATE无法被赋值，需为[DATE]
	 * 
	 * @param param
	 * @throws Exception
	 */
	public static void listToExcel(Param param) throws Exception {
		long t0 = System.currentTimeMillis();
		String templatePath = param.getTemplateExcepPath();
		if (templatePath == null || "".equals(templatePath)) {
			System.out.println("模板Excel不能为空");
			throw new Exception("模板Excel路径不能为空! getTemplatePath()=null ");
		}

		// 创建Excel文档
		Workbook workbook = null;
		boolean excel07Flag = false;
		try {
			workbook = paresExcel(templatePath);
		} catch (Exception e) {
			System.out.println("打开模板失败!");
			throw new Exception("打开模板Excel失败!templatePath=" + templatePath);
		}

		if (workbook instanceof XSSFWorkbook) {
			excel07Flag = true;
		}

		String fileName = param.getExpExcelName();
		if (fileName == null || "".equals(fileName)) {
			fileName = "导出报表.xls";
		} else {
			fileName = parseExcelName(fileName);
			if (excel07Flag && fileName.endsWith("xls")) {
				fileName += "x";// 如果是07版，未传入生成文件后缀，需要补上
			}
		}

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
			for (Entry<String, Object> entry : dateMap.entrySet())
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

		int startLoopRowNum = -1;
		List<Map<String, Object>> dataList = param.getDataList();
		if (dataList != null && startLoopRow != null) {
			startLoopRowNum = startLoopRow.getRowNum();
			int listSize = 0;
			// 赋值循环变量部分
			listSize = dataList.size();
			int k = startLoopRowNum;
			int startLoopColumNum = startLoopRow.getLastCellNum();
			if (listSize != 0) {
				sheet.shiftRows(startLoopRowNum + 1, countRowNum, listSize);// 先将循环下方内容向后平移listSize的行数，便于后边插入表格数据
			}

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

					Object obj = dataList.get(i).get(cellValue);

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
		if (startLoopRow != null) {
			sheet.removeRow(startLoopRow);
			// 删除变量名所在行，并将下面行上移
			sheet.shiftRows(startLoopRowNum + 1, sheet.getLastRowNum(), -1);
		}

		// 隐藏列
		int[] arr = param.getHiddenColumns();
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
		OutputStream out = param.getOutputStream();
		if (out == null) {
			throw new Exception("输出流参数OutputStream不能为空！");
		}
		workbook.write(out);
		out.flush();
		out.close();
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

	private static Workbook paresExcel(String excelpath) throws Exception {
		InputStream in = new FileInputStream(excelpath);
		return excelpath.endsWith("xlsx") ? new XSSFWorkbook(in) : new HSSFWorkbook(new POIFSFileSystem(in));
	}
}