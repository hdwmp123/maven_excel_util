package excel.read;

import java.io.InputStream;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import org.apache.commons.collections.MapUtils;
import org.apache.commons.lang.StringUtils;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.xml.sax.Attributes;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;
import org.xml.sax.helpers.XMLReaderFactory;

import excel.utils.PropertiesUtil;
import excel.utils.Utils;

/**
 * 抽象Excel2007读取器，excel2007的底层数据结构是xml文件，采用SAX的事件驱动的方法解析
 * xml，需要继承DefaultHandler，在遇到文件内容时，事件会触发，这种做法可以大大降低 内存的耗费，特别使用于大数据量的文件。
 */
public class Excel2007Reader extends DefaultHandler {
	private static transient final Logger LOGGER = LoggerFactory
			.getLogger(Excel2007Reader.class);

	@Override
	public void endDocument() throws SAXException {
		super.endDocument();
	}

	private SharedStringsTable sst;// 共享字符串表
	private String lastContents;// 上一次的内容
	private int sheetIndex = -1;// 工作表索引
	private List<Object> rowList = new ArrayList<Object>();// 行集合
	private int curRow = 0;// 当前行
	private int curCol = 0;// 当前列索引
	private int preCol = 0;// 上一列列索引
	private int titleRow = 0;// 标题行，一般情况下为0
	private int dataRow = 0;// 数据行，一般情况下为0
	private int colSize = 0; // 列数
	private IRowReader rowReader;// Excel数据逻辑处理
	private boolean isTElement;// T元素标识
	private CellDataType nextDataType = CellDataType.SSTINDEX;// 单元格数据类型，默认为字符串类型
	private final DataFormatter formatter = new DataFormatter();
	private short formatIndex;
	private String formatString;
	/**
	 * 单元格
	 */
	private StylesTable stylesTable;

	private SimpleDateFormat simpleDateFormat = null;

	private Map<String, Object> returnResult = null;
	private boolean needStop = false;// 是否需要停止解析

	// #############################################################################################
	// get set
	public void setRowReader(IRowReader rowReader) {
		this.rowReader = rowReader;
	}

	public void setDateFormat(String dateFormat) {
		if (StringUtils.isNotBlank(dateFormat)) {
			this.simpleDateFormat = new SimpleDateFormat(dateFormat);
		}
	}

	public int getTitleRow() {
		return this.titleRow;
	}

	public void setTitleRow(int titleRow) {
		this.titleRow = titleRow;
	}

	public int getDataRow() {
		return this.dataRow;
	}

	public void setDataRow(int dataRow) {
		this.dataRow = dataRow;
	}

	// #############################################################################################

	/**
	 * 只遍历一个sheet，其中sheetId为要遍历的sheet索引，从1开始，1-3
	 * 
	 * @param filePath
	 *            文件路径
	 * @param sheetId
	 * @throws Exception
	 */
	public void processOneSheet(String filePath, int sheetId) throws Exception {
		this.execute(filePath, null, sheetId);
	}

	/**
	 * 只遍历一个sheet，其中sheetId为要遍历的sheet索引，从1开始，1-3
	 * 
	 * @param inputStream
	 *            文件流
	 * @param sheetId
	 * @throws Exception
	 */
	public void processOneSheet(InputStream inputStream, int sheetId)
			throws Exception {
		this.execute(null, inputStream, sheetId);
	}

	/**
	 * 遍历 excel 文件
	 * 
	 * @param filePath
	 *            文件路径
	 * @throws Exception
	 */
	public void process(String filePath) throws Exception {
		this.execute(filePath, null, -1);
	}

	/**
	 * 遍历 excel 文件
	 * 
	 * @param inputStream
	 * @throws Exception
	 */
	public void process(InputStream inputStream) throws Exception {
		this.execute(null, inputStream, -1);
	}

	private void execute(String filePath, InputStream inputStream, int sheetId)
			throws Exception {
		OPCPackage pkg = null;
		if (StringUtils.isNotBlank(filePath)) {
			pkg = OPCPackage.open(filePath);
		} else {
			pkg = OPCPackage.open(inputStream);
		}
		XSSFReader r = new XSSFReader(pkg);
		SharedStringsTable sst = r.getSharedStringsTable();
		this.stylesTable = r.getStylesTable();
		XMLReader parser = fetchSheetParser(sst);
		InputStream sheet = null;
		InputSource sheetSource = null;
		if (sheetId > 0) {
			// 根据 rId# 或 rSheet# 查找sheet
			sheet = r.getSheet("rId" + sheetId);
			this.sheetIndex++;
			sheetSource = new InputSource(sheet);
			parser.parse(sheetSource);
			sheet.close();
		} else {
			Iterator<InputStream> sheets = r.getSheetsData();
			while (sheets.hasNext()) {
				this.curRow = 0;
				this.sheetIndex++;
				sheet = sheets.next();
				sheetSource = new InputSource(sheet);
				parser.parse(sheetSource);
				sheet.close();
				break;
			}
		}
		pkg.close();
	}

	public XMLReader fetchSheetParser(SharedStringsTable sst)
			throws SAXException {
		// XMLReader parser =
		// XMLReaderFactory.createXMLReader("org.apache.xerces.parsers.SAXParser");
		XMLReader parser = XMLReaderFactory.createXMLReader();
		this.sst = sst;
		parser.setContentHandler(this);
		return parser;
	}

	@Override
	public void startElement(String uri, String localName, String name,
			Attributes attributes) throws SAXException {
		if (this.needStop) {
			return;
		}
		// c => 单元格
		if ("c".equals(name)) {
			String rowStr = attributes.getValue("r");
			this.curCol = this.getRowIndex(rowStr);
			// 设定单元格类型
			this.setNextDataType(attributes);
		}
		// 当元素为t时
		if ("t".equals(name)) {
			this.isTElement = true;
		} else {
			this.isTElement = false;
		}
		// 置空
		this.lastContents = "";
	}

	@Override
	public void endElement(String uri, String localName, String name)
			throws SAXException {
		if (this.needStop) {
			if (name.equals("row")) {
				if (this.curRow != 0
						&& this.curRow
								% PropertiesUtil.LIMIT_SIZE == 0) {
					LOGGER.info(String.format("回调函命令数停止解析数据-2007,行:%s",
							this.curRow));
				}
				this.curRow++;
			}
			return;
		}
		// t元素也包含字符串
		if (this.isTElement) {
			// 将单元格内容加入rowlist中，在这之前先去掉字符串前后的空白符
			String value = this.lastContents.trim();
			for (int i = this.rowList.size(); i < this.curCol - 1; i++) {
				this.rowList.add(i, null);
			}
			this.isTElement = false;
			this.preCol = this.curCol;
			this.rowList.add(this.curCol - 1, value);
		} else if (name.equals("v")) {
			// v => 单元格的值，如果单元格是字符串则v标签的值为该字符串在SST中的索引
			Object value = this.lastContents.trim();
			for (int i = this.rowList.size(); i < this.curCol - 1; i++) {
				this.rowList.add(i, null);
			}
			value = (value == null ? null : this.getDataValue(value.toString()
					.trim()));
			this.preCol = this.curCol;
            if (this.curCol > 0) {
			    this.rowList.add(this.curCol - 1, value);
			}
		} else {
			// 如果标签名称为 row ，这说明已到行尾，调用 optRows() 方法
			if (name.equals("row")) {
				int tmpCols = this.rowList.size();
				if (this.curRow >= this.dataRow && tmpCols < this.colSize) {
					for (int i = this.rowList.size(); i < this.colSize; i++) {
						this.rowList.add(i, null);
					}
				}
				if (this.curRow == this.titleRow || this.curRow >= this.dataRow) {
					if (this.colSize > 0 && this.rowList.size() > this.colSize) {
						int size = this.rowList.size();
						for (int i = size - 1; i >= this.colSize; i--) {
							rowList.remove(i);
						}
					}
					this.returnResult = this.rowReader.getRows(this.sheetIndex,
							this.curRow, this.rowList);
				}
				if (this.curRow != 0
						&& this.curRow
								% PropertiesUtil.LIMIT_SIZE == 0) {
					LOGGER.info(String.format("解析数据-2007,行:%s", this.curRow));
				}
				if (this.returnResult != null) {
					this.needStop = MapUtils.getBooleanValue(this.returnResult,
							"needStop");
				}
				if (this.curRow == this.titleRow) {
					this.colSize = this.rowList.size();
				}
				this.rowList.clear();
				this.curRow++;
				this.curCol = 0;
				this.preCol = 0;
			}
		}
	}

	@Override
	public void characters(char[] ch, int start, int length)
			throws SAXException {
		if (this.needStop) {
			return;
		}
		// 得到单元格内容的值
		this.lastContents += new String(ch, start, length);
	}

	// 得到列索引，每一列c元素的r属性构成为字母加数字的形式，字母组合为列索引，数字组合为行索引，
	// 如AB45,表示为第（A-A+1）*26+（B-A+1）*26列，45行
	public int getRowIndex(String rowStr) {
		rowStr = rowStr.replaceAll("[^A-Z]", "");
		byte[] rowAbc = rowStr.getBytes();
		int len = rowAbc.length;
		float num = 0;
		for (int i = 0; i < len; i++) {
			num += (rowAbc[i] - 'A' + 1) * Math.pow(26, len - i - 1);
		}
		return (int) num;
	}

	/**
	 * 单元格中的数据可能的数据类型
	 */
	enum CellDataType {
		BOOL, ERROR, FORMULA, INLINESTR, SSTINDEX, NUMBER, DATE, NULL
	}

	/**
	 * 处理数据类型
	 * 
	 * @param attributes
	 */
	private void setNextDataType(Attributes attributes) {
		if (this.curRow == this.titleRow) {
			this.nextDataType = CellDataType.SSTINDEX;
			return;
		}
		this.nextDataType = CellDataType.NUMBER;
		this.formatIndex = -1;
		this.formatString = null;
		String cellType = attributes.getValue("t");
		String cellStyleStr = attributes.getValue("s");

		if ("b".equals(cellType)) {
			this.nextDataType = CellDataType.BOOL;
		} else if ("e".equals(cellType)) {
			this.nextDataType = CellDataType.ERROR;
		} else if ("inlineStr".equals(cellType)) {
			this.nextDataType = CellDataType.INLINESTR;
		} else if ("s".equals(cellType)) {
			this.nextDataType = CellDataType.SSTINDEX;
		} else if ("str".equals(cellType)) {
			this.nextDataType = CellDataType.FORMULA;
		}

		if (cellStyleStr != null) {
			int styleIndex = Integer.parseInt(cellStyleStr);
			XSSFCellStyle style = this.stylesTable.getStyleAt(styleIndex);
			this.formatIndex = style.getDataFormat();
			this.formatString = style.getDataFormatString();

			if ("m/d/yy" == this.formatString) {
				this.nextDataType = CellDataType.DATE;
				this.formatString = "yyyy-MM-dd hh:mm:ss.SSS";
			}

			if (this.formatString == null) {
				this.nextDataType = CellDataType.NULL;
				this.formatString = BuiltinFormats
						.getBuiltinFormat(this.formatIndex);
			}
		}
	}

	/**
	 * 对解析出来的数据进行类型处理
	 * 
	 * @param value
	 *            单元格的值（这时候是一串数字）
	 * @return
	 */
	private Object getDataValue(String value) {
		if (!StringUtils.isNotBlank(value)) {
			return null;
		}
		Object returnValue = null;
		String tempStr = null;
		switch (this.nextDataType) {
		// 这几个的顺序不能随便交换，交换了很可能会导致数据错误
		case BOOL:
			char first = value.charAt(0);
			returnValue = (first == '0' ? false : true);
			break;
		case ERROR:
			returnValue = "\"ERROR:" + value.toString() + '"';
			break;
		case FORMULA:
			returnValue = '"' + value.toString() + '"';
			break;
		case INLINESTR:
			XSSFRichTextString rtsi = new XSSFRichTextString(value.toString());
			returnValue = rtsi.toString();
			rtsi = null;
			break;
		case SSTINDEX:
			String sstIndex = value.toString();
			try {
				int idx = Integer.parseInt(sstIndex);
				XSSFRichTextString rtss = new XSSFRichTextString(
						this.sst.getEntryAt(idx));
				returnValue = rtss.toString();
				rtss = null;
			} catch (NumberFormatException ex) {
				returnValue = value.toString();
			}
			break;
		case NUMBER:
			if (this.formatString != null) {
				returnValue = this.formatter.formatRawCellContents(
						Double.parseDouble(value), this.formatIndex,
						this.formatString).trim();
			} else {
				returnValue = value;
			}
			returnValue = returnValue.toString().replace("_", "").trim();
			break;
		case DATE:
			tempStr = this.formatter.formatRawCellContents(
					Double.parseDouble(value), this.formatIndex,
					this.formatString);
			// 对日期字符串作特殊处理
			tempStr = tempStr.replace("T", " ");
			try {
				returnValue = Utils.parseStringToDate(tempStr,
						new String[] { "yyyy-MM-dd", "yyyy-MM-dd HH:mm:ss",
								"yyyy/MM/dd", "yyyy/MM/dd HH:mm:ss",
								"yyyy.MM.dd", "yyyy.MM.dd HH:mm:ss",
								"yyyyMMdd", "yyyyMMdd HH:mm:ss" });
			} catch (ParseException e) {
				returnValue = new Date(Long.valueOf(value));
			}
			break;
		default:
			returnValue = "";
			break;
		}
		return returnValue;
	}
}
