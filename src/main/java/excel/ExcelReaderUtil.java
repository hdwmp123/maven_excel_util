package excel;

import java.io.InputStream;

import org.apache.commons.lang.StringUtils;

import excel.read.Excel2003Reader;
import excel.read.Excel2007Reader;
import excel.read.IRowReader;
import excel.utils.GlobalContext;

public class ExcelReaderUtil {

	/**
	 * 读取Excel文件，可能是03也可能是07版本
	 * 
	 * @param reader
	 *            数据解析器
	 * @param fileExt
	 *            文件扩展名
	 * @param filePath
	 *            文件路径 与inputStream互斥
	 * @param inputStream
	 *            数据流与filePath互斥
	 * @param sheetId
	 *            sheet索引 从1开始(只对07版本有用)
	 * @param titleRow
	 *            标题行，一般情况下为0 (只对07版本有用)
	 * @param dateFormat
	 *            日期格式 (只对03版有用)
	 * @throws Exception
	 */
	public void readExcel(IRowReader reader, String fileExt, String filePath,
			InputStream inputStream, int sheetId, int titleRow, int dataRow,
			String dateFormat) throws Exception {

		if (fileExt.equals(GlobalContext.EXCEL03_EXTENSION)) {// 处理excel2003文件
			Excel2003Reader excel03 = new Excel2003Reader();
			excel03.setRowReader(reader);
			excel03.setDateFormat(dateFormat);
			excel03.setTitleRow(titleRow);
			excel03.setDataRow(dataRow);
			if (StringUtils.isNotBlank(filePath)) {
				excel03.process(filePath);
			} else {
				excel03.process(inputStream);
			}
		} else if (fileExt.equals(GlobalContext.EXCEL07_EXTENSION)) { // 处理excel2007文件
			Excel2007Reader excel07 = new Excel2007Reader();
			excel07.setRowReader(reader);
			excel07.setDateFormat(dateFormat);
			excel07.setTitleRow(titleRow);
			excel07.setDataRow(dataRow);
			if (StringUtils.isNotBlank(filePath) && sheetId > 0) {
				excel07.processOneSheet(filePath, sheetId);
			} else if (StringUtils.isNotBlank(filePath)) {
				excel07.process(filePath);
			} else if (inputStream != null && sheetId > 0) {
				excel07.processOneSheet(inputStream, sheetId);
			} else if (inputStream != null) {
				excel07.process(inputStream);
			}
		} else {
			throw new Exception("文件格式错误，fileName的扩展名只能是xls或xlsx。");
		}
	}
}
