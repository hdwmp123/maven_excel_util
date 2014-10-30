package excel.utils;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;
import java.text.DateFormat;
import java.text.DecimalFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;
import java.util.Locale;
import java.util.TimeZone;
import java.util.regex.Pattern;

import org.apache.commons.lang.StringUtils;
import org.apache.commons.lang.math.RandomUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.util.IOUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import excel.read.Excel2003Reader;

public class Utils {

	private static transient final Logger logger = LoggerFactory
			.getLogger(Excel2003Reader.class);

	public static String dateToString(Date date, String formate) {
		String ret = "";
		if (date != null) {
			SimpleDateFormat dateFormat = new SimpleDateFormat(formate);
			dateFormat.setTimeZone(TimeZone.getTimeZone("Asia/Shanghai"));
			ret = dateFormat.format(date);
		}

		return ret;
	}

	public static Date stringToDate(String dateString, String formate)
			throws ParseException {
		Date date = new SimpleDateFormat(formate, Locale.CHINA)
				.parse(dateString);
		return date;
	}

	public static Date getDateAdd(Date date, int dayNum) {
		Calendar cal = Calendar.getInstance();
		cal.setTime(date);
		int dayOfYear = cal.get(Calendar.DAY_OF_YEAR);
		cal.set(Calendar.DAY_OF_YEAR, dayOfYear + dayNum);
		return cal.getTime();
	}

	public static int getDayDiff(Date littleDate, Date bigDate) {
		long bigTime = bigDate.getTime();
		long littleTime = littleDate.getTime();
		Double diff = Math.ceil((bigTime - littleTime)
				/ (double) (1000 * 60 * 60 * 24));
		return diff.intValue();
	}

	public static Date parseStringToDate(String dateStr, String formate)
			throws ParseException {
		DateFormat dateFormat = new SimpleDateFormat(formate);
		if (dateStr == null || dateStr.trim().equals("")) {
			return null;
		}
		dateStr = dateStr.trim();
		Date date = null;
		date = dateFormat.parse(dateStr);
		return date;
	}

	public static Date parseStringToDate(String dateStr, String formates[])
			throws ParseException {
		Date date = null;
		for (int i = 0; i < formates.length; i++) {
			String formate = formates[i];
			try {
				date = parseStringToDate(dateStr, formate);
				break;
			} catch (Exception e) {
				continue;
			}
		}
		return date;
	}

	public static double formatNumber(double number, int pointAfterNum) {
		if (pointAfterNum < 0) {
			pointAfterNum = 0;
		}
		StringBuffer str = new StringBuffer("0");
		if (pointAfterNum > 0) {
			str.append(".");
			for (int i = 0; i < pointAfterNum; i++) {
				str.append("#");
			}
		}
		DecimalFormat format = new DecimalFormat(str.toString());
		String formatNumber = format.format(number);
		return Double.parseDouble(formatNumber);
	}

	/**
	 * 保留小数显示数字
	 * 
	 * @param number
	 * @param pointAfterNum
	 *            小数点后位数
	 * @return
	 */
	public static BigDecimal formatNumber(BigDecimal number, int pointAfterNum) {
		if (number == null) {
			return new BigDecimal(0);
		}
		double formatNumber = formatNumber(number.doubleValue(), pointAfterNum);
		return new BigDecimal(formatNumber);
	}

	/**
	 * 获取单元格的值
	 * 
	 * @param cell
	 * @return
	 */
	public static String getCellValue(Cell cell) {
		if (null == cell) {
			return null;
		}
		int cellType = cell.getCellType();
		if (cellType == Cell.CELL_TYPE_BLANK) {
			return null;
		} else if (cellType == Cell.CELL_TYPE_BOOLEAN) {
			boolean cellValue = cell.getBooleanCellValue();
			return String.valueOf(cellValue);
		} else if (cellType == Cell.CELL_TYPE_NUMERIC) {
			if (DateUtil.isCellDateFormatted(cell)) {
				Date date = cell.getDateCellValue();
				if (null == date) {
					return null;
				}
				return dateToString(date, "yyyy-MM-dd HH:mm:ss");
			} else {
				/*
				 * Double cellValue = cell.getNumericCellValue(); if (cellValue
				 * == cellValue.longValue()) {// 整数 return
				 * String.valueOf(cellValue.longValue()); } else { String value
				 * = String.valueOf(cellValue); if (value.indexOf("E") > -1) {//
				 * 科学计数法 value = new BigDecimal(value).toPlainString(); } return
				 * value; }
				 */
				Double cellValue = cell.getNumericCellValue();
				DecimalFormat df = new DecimalFormat("0.#####");
				return df.format(cellValue);
			}
		} else if (cellType == Cell.CELL_TYPE_STRING) {
			return cell.getStringCellValue();
		} else if (cellType == Cell.CELL_TYPE_FORMULA) {// 公式类型
			if (cell.getCachedFormulaResultType() == Cell.CELL_TYPE_NUMERIC) {
				if (DateUtil.isCellDateFormatted(cell)) {
					return dateToString(cell.getDateCellValue(),
							"yyyy-MM-dd HH:mm:ss");
				} else {
					/*
					 * Double cellVal = cell.getNumericCellValue(); String value
					 * = null; if (cellVal == cellVal.longValue()) { value =
					 * String.valueOf(cellVal.longValue()); } else { value =
					 * String.valueOf(cellVal); if (value.indexOf("E") > -1) {
					 * value = new BigDecimal(value).toPlainString(); } } return
					 * value;
					 */
					Double cellValue = cell.getNumericCellValue();
					DecimalFormat df = new DecimalFormat("0.#####");
					return df.format(cellValue);
				}
			} else if (cell.getCachedFormulaResultType() == Cell.CELL_TYPE_STRING) {
				return cell.getRichStringCellValue().toString();
			} else {
				return null;
			}
		} else {
			return null;
		}
	}

	/**
	 * 导出excel数据
	 * 
	 * @param title
	 *            sheet名称
	 * @param headers
	 *            标题行
	 * @param data
	 *            数据
	 * @return
	 * @throws IOException
	 */
	public static byte[] exportExcel(String title, String[] headers,
			String[][] data) throws IOException {
		Workbook book = new HSSFWorkbook();
		Sheet sheet = book.createSheet(title);
		Row headerRow = sheet.createRow(0);
		for (int i = 0; i < headers.length; i++) {
			headerRow.createCell(i, Cell.CELL_TYPE_STRING).setCellValue(
					headers[i]);
		}
		for (int j = 0; j < data.length; j++) {
			Row dataRow = sheet.createRow(1 + j);
			String[] dataLine = data[j];
			for (int k = 0; k < dataLine.length; k++) {
				dataRow.createCell(k, Cell.CELL_TYPE_STRING).setCellValue(
						dataLine[k]);
			}
		}
		ByteArrayOutputStream byteOut = new ByteArrayOutputStream();
		book.write(byteOut);
		return byteOut.toByteArray();
	}

	/**
	 * @Description 导出csv文件
	 * @param exportData
	 * @param title
	 * @return
	 */
	public static String createCSVData(List<String[]> exportData, String[] title) {
		if (exportData == null || exportData.size() == 0 || null == title
				|| title.length == 0) {
			logger.info("组织csv数据缺少必须参数,exportData=" + exportData + ",title="
					+ title);
			return null;
		}
		StringBuilder result = new StringBuilder();
		for (int i = 0; i < title.length; i++) {
			result.append(title[i]);
			if (i == title.length - 1) {
				result.append("\r\n");
			} else {
				result.append(",");
			}
		}
		String[] data = null;
		for (int i = 0; i < exportData.size(); i++) {
			data = (String[]) exportData.get(i);
			if (null == data) {
				continue;
			}
			for (int j = 0; j < data.length; j++) {
				result.append(data[j]);
				if (j == title.length - 1) {
					result.append("\r\n");
				} else {
					result.append(",");
				}
			}
		}
		return result.toString();
	}

	/**
	 * 过滤HTML敏感字符
	 * 
	 * @param value
	 * @return
	 */
	public static String scriptingFilter(String value) {
		if (value == null) {
			return null;
		}
		StringBuffer result = new StringBuffer(value.length());
		for (int i = 0; i < value.length(); ++i) {
			switch (value.charAt(i)) {
			case '<':
				result.append("&lt;");
				break;
			case '>':
				result.append("&gt;");
				break;
			case '"':
				result.append("&quot;");
				break;
			case '\'':
				result.append("&#39;");
				break;
			case '%':
				result.append("&#37;");
				break;
			case ';':
				result.append("&#59;");
				break;
			case '(':
				result.append("&#40;");
				break;
			case ')':
				result.append("&#41;");
				break;
			case '&':
				result.append("&amp;");
				break;
			case '+':
				result.append("&#43;");
				break;
			default:
				result.append(value.charAt(i));
				break;
			}
		}
		return new String(result);
	}

	/**
	 * 返回aList不包含在bList中的元素集合
	 * 
	 * @param aList
	 * @param bList
	 * @return
	 */
	public static List<Long> notContainedList(List<Long> aList, List<Long> bList) {
		if (aList == null || bList == null) {
			return null;
		}
		List<Long> resultList = new ArrayList<Long>();
		for (Long a : aList) {
			boolean flag = false;
			for (Long b : bList) {
				if (a.longValue() == b.longValue()) {
					flag = true;
					break;
				}
			}
			if (!flag) {
				resultList.add(a);
			}
		}
		return resultList;
	}

	/**
	 * 返回aList包含在bList中的元素集合
	 * 
	 * @param aList
	 * @param bList
	 * @return
	 */
	public static List<Long> containedList(List<Long> aList, List<Long> bList) {
		if (aList == null || bList == null) {
			return null;
		}
		List<Long> resultList = new ArrayList<Long>();
		for (Long a : aList) {
			for (Long b : bList) {
				if (a.longValue() == b.longValue()) {
					resultList.add(a);
					break;
				}
			}
		}
		return resultList;
	}

	/**
	 * 获取文件大小
	 * 
	 * @param inputStream
	 *            文件输入流
	 * @param unit
	 *            单位：B,KB,MB,GB
	 * @return
	 */
	public static double getFileSize(InputStream inputStream, String unit) {
		double fileSize = 0;
		if (unit == null) {
			return fileSize;
		}
		try {
			int size = inputStream.available();
			if ("B".equals(unit.toUpperCase())) {
				fileSize = size;
			} else if ("KB".equals(unit.toUpperCase())) {
				fileSize = size / (double) 1024;
			} else if ("MB".equals(unit.toUpperCase())) {
				fileSize = size / ((double) 1024 * 1024);
			} else if ("GB".equals(unit.toUpperCase())) {
				fileSize = size / ((double) 1024 * 1024 * 1024);
			}
		} catch (Exception e) {
			logger.error("获取文件大小异常：" + e.getMessage());
		}
		return fileSize;
	}

	/**
	 * 创建临时文件
	 * 
	 * @param filePath
	 *            文件路径（包含文件名）
	 * @return 临时文件名
	 * @throws Exception
	 */
	public static String createTempFile(String filePath) throws Exception {
		File file = new File(filePath);
		if (file.exists() && file.isFile()) {
			String fileName = file.getName();
			String tempFileName = null;
			if (fileName.indexOf(".") != -1) {
				String fileNameNoExt = fileName.substring(0,
						fileName.lastIndexOf("."));
				String fileExt = fileName
						.substring(fileName.lastIndexOf(".") + 1);
				tempFileName = fileNameNoExt + RandomUtils.nextInt() + "."
						+ fileExt;
			} else {
				tempFileName = fileName + RandomUtils.nextInt();
			}
			FileInputStream input = new FileInputStream(file);
			FileOutputStream output = new FileOutputStream(new File(
					file.getParent() + File.separator + tempFileName));
			IOUtils.copy(input, output);
			output.flush();
			output.close();
			input.close();
			return tempFileName;
		}
		return null;
	}

	/**
	 * 将骆驼命名法字段转换成数据库字段，如userId -> user_id
	 * 
	 * @param column
	 * @return
	 */
	public static String convertToDbColumn(String column) {
		if (StringUtils.isBlank(column)) {
			return null;
		}
		StringBuilder result = new StringBuilder();
		int size = column.length();
		for (int i = 0; i < size; i++) {
			char columnChar = column.charAt(i);
			if (columnChar >= 'A' && columnChar <= 'Z') {
				char newColumnChar = (char) (columnChar + 32);
				result.append('_').append(newColumnChar);
			} else {
				result.append(columnChar);
			}
		}
		return result.toString();
	}

	/**
	 * 根据日期字符串值获取对应的日期格式，如传入2011-01-01，返回yyyy-MM-dd <br/>
	 * 日期可选格式为： <br/>
	 * "yyyy-MM-dd" <br/>
	 * "yyyy/MM/dd" <br/>
	 * "yyyy.MM.dd" <br/>
	 * "yyyyMMdd" <br/>
	 * "yyyy-MM-dd HH:mm:ss" <br/>
	 * "yyyy/MM/dd HH:mm:ss" <br/>
	 * "yyyy.MM.dd HH:mm:ss" <br/>
	 * "yyyyMMdd HH:mm:ss"
	 * 
	 * @param date
	 *            日期字符串值
	 * @return 如果传入值可转化为日期，返回相应格式，否则返回null
	 */
	public static String validateDate(String date) {
		String result = null;
		for (int i = 0; i < patterns.size(); i++) {
			if (patterns.get(i).matcher(date).matches()) {
				try {
					Date formatValue = new SimpleDateFormat(formats[i])
							.parse(date);
					if (formatValue != null) {
						result = formats[i];
					}
				} catch (ParseException e) {
					result = null;
				}
				break;
			}
		}
		return result;
	}

	private static String[] regexs = { "\\d{4}\\-\\d{1,2}\\-\\d{1,2}",
			"\\d{4}\\/\\d{1,2}\\/\\d{1,2}", "\\d{4}\\.\\d{1,2}\\.\\d{1,2}",
			"\\d{4}\\d{1,2}\\d{1,2}",
			"\\d{4}\\-\\d{1,2}\\-\\d{1,2}[ ]\\d{1,2}\\:\\d{1,2}\\:\\d{1,2}",
			"\\d{4}\\/\\d{1,2}\\/\\d{1,2}[ ]\\d{1,2}\\:\\d{1,2}\\:\\d{1,2}",
			"\\d{4}\\.\\d{1,2}\\.\\d{1,2}[ ]\\d{1,2}\\:\\d{1,2}\\:\\d{1,2}",
			"\\d{4}\\d{1,2}\\d{1,2}[ ]\\d{1,2}\\:\\d{1,2}\\:\\d{1,2}" };
	private static String[] formats = { "yyyy-MM-dd", "yyyy/MM/dd",
			"yyyy.MM.dd", "yyyyMMdd", "yyyy-MM-dd HH:mm:ss",
			"yyyy/MM/dd HH:mm:ss", "yyyy.MM.dd HH:mm:ss", "yyyyMMdd HH:mm:ss" };
	private static List<Pattern> patterns = new ArrayList<Pattern>();
	static {
		for (int i = 0; i < regexs.length; i++) {
			patterns.add(Pattern.compile(regexs[i]));
		}
	}
}
