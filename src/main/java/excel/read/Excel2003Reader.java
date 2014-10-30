package excel.read;

import java.io.FileInputStream;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.Map;

import org.apache.commons.collections.MapUtils;
import org.apache.commons.lang.StringUtils;
import org.apache.poi.hssf.eventusermodel.EventWorkbookBuilder.SheetRecordCollectingListener;
import org.apache.poi.hssf.eventusermodel.FormatTrackingHSSFListener;
import org.apache.poi.hssf.eventusermodel.HSSFEventFactory;
import org.apache.poi.hssf.eventusermodel.HSSFListener;
import org.apache.poi.hssf.eventusermodel.HSSFRequest;
import org.apache.poi.hssf.eventusermodel.MissingRecordAwareHSSFListener;
import org.apache.poi.hssf.eventusermodel.dummyrecord.LastCellOfRowDummyRecord;
import org.apache.poi.hssf.eventusermodel.dummyrecord.MissingCellDummyRecord;
import org.apache.poi.hssf.model.HSSFFormulaParser;
import org.apache.poi.hssf.record.BOFRecord;
import org.apache.poi.hssf.record.BlankRecord;
import org.apache.poi.hssf.record.BoolErrRecord;
import org.apache.poi.hssf.record.BoundSheetRecord;
import org.apache.poi.hssf.record.FormulaRecord;
import org.apache.poi.hssf.record.LabelRecord;
import org.apache.poi.hssf.record.LabelSSTRecord;
import org.apache.poi.hssf.record.NoteRecord;
import org.apache.poi.hssf.record.NumberRecord;
import org.apache.poi.hssf.record.RKRecord;
import org.apache.poi.hssf.record.Record;
import org.apache.poi.hssf.record.SSTRecord;
import org.apache.poi.hssf.record.StringRecord;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.DateUtil;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import excel.utils.PropertiesUtil;

/**
 * 抽象Excel2003读取器，通过实现HSSFListener监听器，采用事件驱动模式解析excel2003
 * 中的内容，遇到特定事件才会触发，大大减少了内存的使用。
 * 
 */
public class Excel2003Reader implements HSSFListener {

	private static transient final Logger LOGGER = LoggerFactory
			.getLogger(Excel2003Reader.class);

	private POIFSFileSystem fs;
	private SheetRecordCollectingListener workbookBuildingListener;
	private HSSFWorkbook stubWorkbook;
	private SSTRecord sstRecord;
	private FormatTrackingHSSFListener formatListener;
	private BoundSheetRecord[] orderedBSRs;
	private SimpleDateFormat simpleDateFormat = null;// excel中日期保存格式
	private String sheetName;

	private List<Object> rowList = new ArrayList<Object>();
	private int curRow = 0;// 当前行
	private int preRow;
	private int preCol;
	private int nextRow;
	private int nextColumn;
	private int minColumns = -1;
	private int sheetIndex = -1;
	private int titleRow = 0;// 标题行，一般情况下为0
	private int dataRow = 0;// 数据行，一般情况下为0
	private int colSize = 0; // 列数

	private ArrayList boundSheetRecords = new ArrayList();
	private boolean outputFormulaValues = true;
	private boolean outputNextStringRecord;

	private IRowReader rowReader;
	private Map<String, Object> returnResult = null;
	private boolean needStop = false;// 是否需要停止解析

	// #############################################################################################
	// get set
	public void setRowReader(IRowReader rowReader) {
		this.rowReader = rowReader;
		this.simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
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
	 * 遍历 excel 文件
	 * 
	 * @param filePath
	 *            文件路径
	 * @throws Exception
	 */
	public void process(String filePath) throws Exception {
		this.execute(filePath, null);
	}

	/**
	 * 遍历 excel 文件
	 * 
	 * @param inputStream
	 * @throws Exception
	 */
	public void process(InputStream inputStream) throws Exception {
		this.execute(null, inputStream);
	}

	private void execute(String filePath, InputStream inputStream)
			throws Exception {

		this.fs = null;
		if (StringUtils.isNotBlank(filePath)) {
			this.fs = new POIFSFileSystem(new FileInputStream(filePath));
		} else {
			this.fs = new POIFSFileSystem(inputStream);
		}

		MissingRecordAwareHSSFListener listener = new MissingRecordAwareHSSFListener(
				this);
		this.formatListener = new FormatTrackingHSSFListener(listener);
		HSSFEventFactory factory = new HSSFEventFactory();
		HSSFRequest request = new HSSFRequest();
		if (this.outputFormulaValues) {
			request.addListenerForAllRecords(this.formatListener);
		} else {
			this.workbookBuildingListener = new SheetRecordCollectingListener(
					this.formatListener);
			request.addListenerForAllRecords(this.workbookBuildingListener);
		}
		factory.processWorkbookEvents(request, this.fs);
	}

	/**
	 * HSSFListener 监听方法，处理 Record
	 */
	@Override
	public void processRecord(Record record) {
		if (this.needStop) {
			if (record instanceof LastCellOfRowDummyRecord) {
				if (this.curRow != 0
						&& this.curRow % this.curRow
								% PropertiesUtil.LIMIT_SIZE == 0) {
					LOGGER.info(String.format("回调函命令数停止解析数据-2003,行:%s",
							this.curRow));
				}
				this.curRow++;
			}
			return;
		}
		int thisRow = -1;
		int thisColumn = -1;
		Object thisStr = null;
		Object value = null;
		switch (record.getSid()) {
		case BoundSheetRecord.sid:
			this.boundSheetRecords.add(record);
			break;
		case BOFRecord.sid:
			BOFRecord br = (BOFRecord) record;
			if (br.getType() == BOFRecord.TYPE_WORKSHEET) {
				// 如果有需要，则建立子工作薄
				if (this.workbookBuildingListener != null
						&& this.stubWorkbook == null) {
					this.stubWorkbook = this.workbookBuildingListener
							.getStubHSSFWorkbook();
				}

				this.sheetIndex++;
				if (this.orderedBSRs == null) {
					this.orderedBSRs = BoundSheetRecord
							.orderByBofPosition(this.boundSheetRecords);
				}
				this.sheetName = this.orderedBSRs[this.sheetIndex]
						.getSheetname();
			}
			break;

		case SSTRecord.sid:
			this.sstRecord = (SSTRecord) record;
			break;

		case BlankRecord.sid:
			BlankRecord brec = (BlankRecord) record;
			thisRow = brec.getRow();
			thisColumn = brec.getColumn();
			thisStr = "";
			this.add(thisColumn, thisStr);
			break;
		case BoolErrRecord.sid: // 单元格为布尔类型
			BoolErrRecord berec = (BoolErrRecord) record;
			thisRow = berec.getRow();
			thisColumn = berec.getColumn();
			thisStr = berec.getBooleanValue() + "";
			this.add(thisColumn, thisStr);
			break;
		case FormulaRecord.sid: // 单元格为公式类型
			FormulaRecord frec = (FormulaRecord) record;
			thisRow = frec.getRow();
			thisColumn = frec.getColumn();
			if (this.outputFormulaValues) {
				if (Double.isNaN(frec.getValue())) {
					// Formula result is a string
					// This is stored in the next record
					this.outputNextStringRecord = true;
					this.nextRow = frec.getRow();
					this.nextColumn = frec.getColumn();
				} else {
					thisStr = this.formatListener.formatNumberDateCell(frec);
				}
			} else {
				thisStr = '"' + HSSFFormulaParser.toFormulaString(
						this.stubWorkbook, frec.getParsedExpression()) + '"';
			}
			this.add(thisColumn, thisStr);
			break;
		case StringRecord.sid:// 单元格中公式的字符串
			if (this.outputNextStringRecord) {
				// String for formula
				StringRecord srec = (StringRecord) record;
				thisStr = srec.getString();
				thisRow = this.nextRow;
				thisColumn = this.nextColumn;
				this.outputNextStringRecord = false;
			}
			break;
		case LabelRecord.sid:
			LabelRecord lrec = (LabelRecord) record;
			this.curRow = thisRow = lrec.getRow();
			thisColumn = lrec.getColumn();
			value = lrec.getValue().trim();
			value = value.equals("") ? "" : value;
			this.add(thisColumn, value);
			break;
		case LabelSSTRecord.sid: // 单元格为字符串类型
			LabelSSTRecord lsrec = (LabelSSTRecord) record;
			this.curRow = thisRow = lsrec.getRow();
			thisColumn = lsrec.getColumn();
			if (this.sstRecord == null) {
				this.add(thisColumn, null);
			} else {
				value = this.sstRecord.getString(lsrec.getSSTIndex())
						.toString().trim();
				value = value.equals("") ? "" : value;
				this.add(thisColumn, value);
			}
			break;
		case NoteRecord.sid:
			NoteRecord nrec = (NoteRecord) record;
			thisRow = nrec.getRow();
			thisColumn = nrec.getColumn();
			break;
		case NumberRecord.sid: // 单元格为数字类型
			NumberRecord numrec = (NumberRecord) record;
			this.curRow = thisRow = numrec.getRow();
			thisColumn = numrec.getColumn();
			// 判断是否是日期格式,如果是日期则返回yyyy-MM-dd HH:mm:ss格式字符串
			double cellVlaue = numrec.getValue();
			if (DateUtil.isADateFormat(
					this.formatListener.getFormatIndex(numrec),
					this.formatListener.getFormatString(numrec))
					&& DateUtil.isValidExcelDate(cellVlaue)) {
				// value =
				// this.simpleDateFormat.format(DateUtil.getJavaDate(cellVlaue));
				value = DateUtil.getJavaDate(cellVlaue);
			} else {
				value = this.formatListener.formatNumberDateCell(numrec).trim();
			}
			value = value.equals("") ? "" : value;
			this.add(thisColumn, value);
			break;
		case RKRecord.sid:
			RKRecord rkrec = (RKRecord) record;
			thisRow = rkrec.getRow();
			thisColumn = rkrec.getColumn();
			break;
		default:
			break;
		}

		// 遇到新行的操作
		if (thisRow != -1 && thisRow != this.preRow) {
			this.preCol = -1;
		}

		// 空值的操作
		if (record instanceof MissingCellDummyRecord) {
			MissingCellDummyRecord mc = (MissingCellDummyRecord) record;
			this.curRow = thisRow = mc.getRow();
			thisColumn = mc.getColumn();
			for (int i = this.rowList.size(); i < thisColumn; i++) {
				this.rowList.add(i, null);
			}
		}

		// 更新行和列的值
		if (thisRow > -1) {
			this.preRow = thisRow;
		}
		if (thisColumn > -1) {
			this.preCol = thisColumn;
		}

		// 行结束时的操作
		if (record instanceof LastCellOfRowDummyRecord) {
			if (this.minColumns > 0) {
				// 列值重新置空
				if (this.preCol == -1) {
					this.preCol = 0;
				}
			}
			// 行结束时， 调用 optRows() 方法
			this.preCol = -1;
			int tmpCols = this.rowList.size();
			if (this.curRow >= this.dataRow && tmpCols < this.colSize) {
				for (int i = 0; i < this.colSize - tmpCols; i++) {
					this.rowList.add(this.rowList.size(), null);
				}
			}
			if (this.curRow == this.titleRow || this.curRow >= this.dataRow) {
				this.returnResult = this.rowReader.getRows(this.sheetIndex,
						this.curRow, this.rowList);
			}
			if (this.curRow != 0
					&& this.curRow
							% PropertiesUtil.LIMIT_SIZE == 0) {
				LOGGER.info(String.format("解析数据-2003,行:%s", this.curRow));
			}
			if (this.returnResult != null) {
				this.needStop = MapUtils.getBooleanValue(this.returnResult,
						"needStop");
			}
			if (this.curRow == this.titleRow) {
				this.colSize = this.rowList.size();
			}
			// 清空容器
			clearRowList();
		}
	}

	/**
	 * 通用list.add 方法
	 * 
	 * @param index
	 * @param value
	 */
	public void add(int index, Object value) {
		if (this.colSize == 0) {
			for (int i = this.rowList.size(); i < index; i++) {
				this.rowList.add(i, null);
			}
			this.rowList.add(index, value);
		} else {
			if (index < this.colSize) {
				this.rowList.set(index, value);
			}
		}
	}

	/**
	 * 清空rowList
	 */
	public void clearRowList() {
		if (this.colSize == 0) {// 当读完标题行后，会把colSize设置为标题行列数
			this.rowList.clear();
		} else {// 当colSize>0时，就确定了rowList内值的个数，全部初始化为null
			Collections.fill(this.rowList, null);
		}
	}
}
