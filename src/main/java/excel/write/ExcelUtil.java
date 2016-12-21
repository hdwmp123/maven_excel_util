package excel.write;

import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStreamWriter;
import java.nio.charset.Charset;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.util.IOUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.csvreader.CsvReader;
import com.csvreader.CsvWriter;

import excel.utils.BeanUtil;
import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.format.CellFormat;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableCellFormat;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

/**
 * excel数据导出
 */
public class ExcelUtil {

	private static final Logger logger = LoggerFactory
			.getLogger(ExcelUtil.class);

	/**
	 * 读取文件中数据
	 * 
	 * @param inputStream
	 *            文件流
	 * @param columnCount
	 *            规定列数
	 * @return @
	 */
	public static List<String[]> getDataFromXls(InputStream inputStream,
			int columnCount) {
		Workbook wb = null;

		// 构造Workbook（工作薄）对象
		try {
			wb = Workbook.getWorkbook(inputStream);
		} catch (BiffException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}

		// 获得了Workbook对象之后，就可以通过它得到Sheet（工作表）对象了
		Sheet[] sheets = wb.getSheets();
		Sheet sheet = (sheets != null && sheets.length > 0) ? sheets[0] : null;

		if (sheet == null) {
			return null;
		}
		if (sheet.getColumns() != columnCount) {// 验证Excel格式
			String msg = "不正确的文件格式,标准格式为" + columnCount + "列";
			logger.info(msg);
			return null;
		}
		//
		List<String[]> datas = new ArrayList<String[]>();
		String[] data = null;
		// 得到当前工作表的行数
		int rowNum = sheet.getRows();
		for (int y = 1; y < rowNum; y++) {// 从第二行开始
			// 得到当前行的所有单元格
			Cell[] cells = sheet.getRow(y);
			data = new String[columnCount];
			for (int x = 0; x < cells.length; x++) {
				data[x] = cells[x].getContents();
			}
			datas.add(data);
		}
		if (wb != null) {
			wb.close();
		}
		//
		return datas;
	}

	/**
	 * 将数据写入excel文件
	 * 
	 * @param sheet
	 *            页
	 * @param title
	 *            标题
	 * @param data
	 *            数据
	 */
	public static void writeDataToXls(WritableSheet sheet, String[] title,
			List<String[]> outputData) {
		Label label = null;
		try {
			CellFormat fm = new WritableCellFormat();
			if (title != null) {
				for (int i = 0; i < title.length; i++) {
					label = new Label(i, 0, title[i], fm);// 第一行标题
					sheet.addCell(label);
				}
			}
			if (outputData != null && outputData.size() > 0) {
				for (int r = 0; r < outputData.size(); r++) {
					String[] row = outputData.get(r);
					for (int c = 0; c < row.length; c++) {
						label = new Label(c, r + 1, row[c], fm);
						sheet.addCell(label);
					}
				}
			}
		} catch (RowsExceededException e) {
			e.printStackTrace();
		} catch (WriteException e) {
			e.printStackTrace();
		}
	}

	/**
	 * 将数据写入excel文件
	 * 
	 * @param file
	 *            文件
	 * @param sheetName
	 *            页名称
	 * @param title
	 *            标题
	 * @param outputData
	 *            数据 @
	 */
	public static void writeDataToXls(File file, String sheetName,
			String[] title, List<String[]> outputData) {
		WritableWorkbook wb = null;
		try {
			wb = Workbook.createWorkbook(file);
			WritableSheet sheet = wb.createSheet(sheetName, 0);// 写入数据
			writeDataToXls(sheet, title, outputData);
		} catch (IOException e) {
			logger.error("生成excel文件失败:" + e.getMessage());
		} finally {
			try {
				if (wb != null) {
					wb.write();
					wb.close();
				}
			} catch (IOException e) {
				e.printStackTrace();
			} catch (WriteException e) {
				e.printStackTrace();
			}
		}
	}

	// #############################################################################################
	/**
	 * 数据写入到csv
	 * 
	 * @param inputStream
	 *            文件流
	 * @param columnCount
	 *            列限制
	 * @return @
	 */
	public static List<String[]> getDataFromCsv(InputStream inputStream,
			int columnCount) {
		CsvReader csvReader = null;
		List<String[]> datas = new ArrayList<String[]>();
		byte[] csvArray;
		try {
			csvArray = IOUtils.toByteArray(inputStream);
			String encoding = BeanUtil.getEncoding(csvArray);
			logger.debug(encoding);
			if (!encoding.equalsIgnoreCase("utf-8")) {
				String unicode = new String(csvArray, "GB2312");
				csvArray = unicode.getBytes("utf-8");
			}
			ByteArrayInputStream bais = new ByteArrayInputStream(csvArray);
			csvReader = new CsvReader(bais, Charset.forName("utf-8"));
			csvReader.readHeaders();
		} catch (IOException e) {
			String msg = "文件解析失败";
			logger.info(msg);
		}

		if (csvReader.getHeaderCount() != columnCount) {// 验证Excel格式
			String msg = "不正确的文件格式,标准格式为" + columnCount + "列";
			logger.info(msg);
		}
		try {
			while (csvReader.readRecord()) {
				String[] record = csvReader.getValues();
				datas.add(record);
			}
		} catch (IOException e) {
			String msg = "文件数据读取失败";
			logger.info(msg);
		}
		if (csvReader != null) {
			csvReader.close();
		}
		return datas;
	}

	/**
	 * 数据写入到csv
	 * 
	 * @param file
	 *            文件
	 * @param title
	 *            标题
	 * @param outputData
	 *            数据 @
	 */
	public static void writeDataToCsv(File file, String[] title,
			List<String[]> outputData) {
		CsvWriter csvWriter = null;
		try {
			OutputStreamWriter out = new OutputStreamWriter(
					new FileOutputStream(file), Charset.forName("GB2312"));
			csvWriter = new CsvWriter(out, ',');
			csvWriter.writeRecord(title);
			for (String[] errorRecord : outputData) {
				csvWriter.writeRecord(errorRecord);
			}
		} catch (IOException e) {
			String msg = "writeDataToCsv数据写入失败";
			logger.info(msg);
		} finally {
			if (csvWriter != null) {
				csvWriter.close();
			}
		}

	}
}
