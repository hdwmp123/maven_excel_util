package excel.read;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.xml.sax.Attributes;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;
import org.xml.sax.helpers.XMLReaderFactory;

public class Excel2007Reader2 extends DefaultHandler {
	/**
	 * 共享字符串表
	 */
	private SharedStringsTable sst;

	/**
	 * 上一次的内容
	 */
	private String lastContents;

	/**
	 * 字符串标识
	 */
	private boolean nextIsString;

	/**
	 * 工作表索引
	 */
	private int sheetIndex = -1;

	/**
	 * 行集合
	 */
	private List<String> rowlist = new ArrayList<String>();

	/**
	 * 当前行
	 */
	private int curRow = 0;

	/**
	 * 当前列
	 */
	private int curCol = 0;

	/**
	 * T元素标识
	 */
	private boolean isTElement;

	/**
	 * Excel数据逻辑处理
	 */
	private IRowReader<String> rowHandler;

	/**
	 * 异常信息，如果为空则表示没有异常
	 */
	private String exceptionMessage;

	/**
	 * 单元格数据类型，默认为字符串类型
	 */
	private CellDataType nextDataType = CellDataType.SSTINDEX;

	private final DataFormatter formatter = new DataFormatter();

	private short formatIndex;

	private String formatString;

	/**
	 * 单元格
	 */
	private StylesTable stylesTable;

	public void setRowReader(IRowReader rowHandler) {
		this.rowHandler = rowHandler;
	}

	/**
	 * 遍历工作簿中所有的电子表格
	 * 
	 * @param filename
	 * @throws IOException
	 * @throws OpenXML4JException
	 * @throws SAXException
	 * @throws Exception
	 */
	public void process(String filename) throws IOException,
			OpenXML4JException, SAXException {
		OPCPackage pkg = OPCPackage.open(filename);
		XSSFReader xssfReader = new XSSFReader(pkg);
		this.stylesTable = xssfReader.getStylesTable();
		SharedStringsTable sst = xssfReader.getSharedStringsTable();
		XMLReader parser = this.fetchSheetParser(sst);
		Iterator<InputStream> sheets = xssfReader.getSheetsData();
		while (sheets.hasNext()) {
			this.curRow = 0;
			this.sheetIndex++;
			InputStream sheet = sheets.next();
			InputSource sheetSource = new InputSource(sheet);
			parser.parse(sheetSource);
			sheet.close();
		}
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
		// c => 单元格
		if ("c".equals(name)) {
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
	public void setNextDataType(Attributes attributes) {
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
	 * @param thisStr
	 *            一个空字符串
	 * @return
	 */
	public String getDataValue(String value, String thisStr) {
		switch (this.nextDataType) {
		// 这几个的顺序不能随便交换，交换了很可能会导致数据错误
		case BOOL:
			char first = value.charAt(0);
			thisStr = first == '0' ? "FALSE" : "TRUE";
			break;
		case ERROR:
			thisStr = "\"ERROR:" + value.toString() + '"';
			break;
		case FORMULA:
			thisStr = '"' + value.toString() + '"';
			break;
		case INLINESTR:
			XSSFRichTextString rtsi = new XSSFRichTextString(value.toString());

			thisStr = rtsi.toString();
			rtsi = null;
			break;
		case SSTINDEX:
			String sstIndex = value.toString();
			try {
				int idx = Integer.parseInt(sstIndex);
				XSSFRichTextString rtss = new XSSFRichTextString(
						this.sst.getEntryAt(idx));
				thisStr = rtss.toString();
				rtss = null;
			} catch (NumberFormatException ex) {
				thisStr = value.toString();
			}
			break;
		case NUMBER:
			if (this.formatString != null) {
				thisStr = this.formatter.formatRawCellContents(
						Double.parseDouble(value), this.formatIndex,
						this.formatString).trim();
			} else {
				thisStr = value;
			}

			thisStr = thisStr.replace("_", "").trim();
			break;
		case DATE:
			thisStr = this.formatter.formatRawCellContents(
					Double.parseDouble(value), this.formatIndex,
					this.formatString);
			// 对日期字符串作特殊处理
			thisStr = thisStr.replace(" ", "T");
			break;
		default:
			thisStr = " ";
			break;
		}

		return thisStr;
	}

	@Override
	public void endElement(String uri, String localName, String name)
			throws SAXException {
		// 根据SST的索引值的到单元格的真正要存储的字符串
		// 这时characters()方法可能会被调用多次
		if (this.nextIsString) {
			int idx = Integer.parseInt(this.lastContents);
			this.lastContents = new XSSFRichTextString(this.sst.getEntryAt(idx))
					.toString();
		}

		// t元素也包含字符串
		if (this.isTElement) {
			// 将单元格内容加入rowlist中，在这之前先去掉字符串前后的空白符
			String value = this.lastContents.trim();
			this.rowlist.add(this.curCol, value);
			this.curCol++;
			this.isTElement = false;
		} else if ("v".equals(name)) {
			// v => 单元格的值，如果单元格是字符串则v标签的值为该字符串在SST中的索引
			String value = this.getDataValue(this.lastContents.trim(), "");

			this.rowlist.add(this.curCol, value);
			this.curCol++;
		} else {
			// 如果标签名称为 row ，这说明已到行尾，调用 optRows() 方法
			if (name.equals("row")) {
				this.rowHandler.getRows(this.sheetIndex, this.curRow,
						this.rowlist);
				this.rowlist.clear();
				this.curRow++;
				this.curCol = 0;
			}
		}
	}

	@Override
	public void characters(char[] ch, int start, int length)
			throws SAXException {
		// 得到单元格内容的值
		this.lastContents += new String(ch, start, length);
	}

	/**
	 * @return the exceptionMessage
	 */
	public String getExceptionMessage() {
		return this.exceptionMessage;
	}
}
