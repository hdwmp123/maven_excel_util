package excel.read;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.sql.SQLException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.eventusermodel.EventWorkbookBuilder.SheetRecordCollectingListener;
import org.apache.poi.hssf.eventusermodel.FormatTrackingHSSFListener;
import org.apache.poi.hssf.eventusermodel.HSSFEventFactory;
import org.apache.poi.hssf.eventusermodel.HSSFListener;
import org.apache.poi.hssf.eventusermodel.HSSFRequest;
import org.apache.poi.hssf.eventusermodel.MissingRecordAwareHSSFListener;
import org.apache.poi.hssf.eventusermodel.dummyrecord.LastCellOfRowDummyRecord;
import org.apache.poi.hssf.eventusermodel.dummyrecord.MissingCellDummyRecord;
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

public abstract class Excel2003Reader2 implements HSSFListener {
	private int minColumns;
	private POIFSFileSystem fs;
	private int lastRowNumber;
	private int lastColumnNumber;

	private boolean outputFormulaValues = true;

	private SheetRecordCollectingListener workbookBuildingListener;
	private HSSFWorkbook stubWorkbook;

	private SSTRecord sstRecord;
	private FormatTrackingHSSFListener formatListener;

	private int sheetIndex = -1;
	private BoundSheetRecord[] orderedBSRs;
	@SuppressWarnings("unchecked")
	private ArrayList boundSheetRecords = new ArrayList();

	private int nextRow;
	private int nextColumn;
	private boolean outputNextStringRecord;

	private int curRow;
	private List<String> rowList;
	private int colSize = 0; // 列数

	// excel中日期保存格式
	private static SimpleDateFormat sdf = new SimpleDateFormat(
			"yyyy-MM-dd HH:mm:ss");

	// 传入参数
	private int titleRow;

	public Excel2003Reader2(POIFSFileSystem fs) throws SQLException {
		this.fs = fs;
		this.minColumns = -1;
		this.curRow = 0;
		this.rowList = new ArrayList<String>();
	}

	public Excel2003Reader2(String filename) throws IOException,
			FileNotFoundException, SQLException {
		this(new POIFSFileSystem(new FileInputStream(filename)));
	}

	// excel记录行操作方法，以行索引和行元素列表为参数，对一行元素进行操作，元素为String类型
	// public abstract void optRows(int curRow, List<String> rowlist) throws
	// SQLException ;

	// excel记录行操作方法，以sheet索引，行索引和行元素列表为参数，对sheet的一行元素进行操作，元素为String类型
	public abstract void optRows(int sheetIndex, int curRow,
			List<String> rowlist) throws SQLException;

	/**
	 * 遍历 excel 文件
	 */
	public void process() throws IOException {
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
	@SuppressWarnings("unchecked")
	public void processRecord(Record record) {
		int thisRow = -1;
		int thisColumn = -1;
		String value = null;
		switch (record.getSid()) {
		case BoundSheetRecord.sid:
			this.boundSheetRecords.add(record);
			break;
		case BOFRecord.sid:
			BOFRecord br = (BOFRecord) record;
			if (br.getType() == BOFRecord.TYPE_WORKSHEET) {
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
			}
			break;

		case SSTRecord.sid:
			this.sstRecord = (SSTRecord) record;
			break;

		case BlankRecord.sid:
			BlankRecord brec = (BlankRecord) record;

			thisRow = brec.getRow();
			thisColumn = brec.getColumn();
			break;
		case BoolErrRecord.sid:
			BoolErrRecord berec = (BoolErrRecord) record;
			thisRow = berec.getRow();
			thisColumn = berec.getColumn();
			break;

		case FormulaRecord.sid:
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
				}
			} else {
			}
			break;
		case StringRecord.sid:
			if (this.outputNextStringRecord) {
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
			value = value.equals("") ? " " : value;
			for (int i = this.rowList.size(); i < thisColumn; i++) {
				this.rowList.add(i, null);
			}
			this.rowList.add(thisColumn, value);
			break;
		case LabelSSTRecord.sid:
			LabelSSTRecord lsrec = (LabelSSTRecord) record;

			this.curRow = thisRow = lsrec.getRow();
			thisColumn = lsrec.getColumn();
			if (this.sstRecord == null) {
				this.rowList.add(thisColumn, null);
			} else {
				value = this.sstRecord.getString(lsrec.getSSTIndex())
						.toString().trim();
				value = value.equals("") ? null : value;
				for (int i = this.rowList.size(); i < thisColumn; i++) {
					this.rowList.add(i, null);
				}
				this.rowList.add(thisColumn, value);
			}
			break;
		case NoteRecord.sid:
			NoteRecord nrec = (NoteRecord) record;
			thisRow = nrec.getRow();
			thisColumn = nrec.getColumn();
			break;
		case NumberRecord.sid:
			NumberRecord numrec = (NumberRecord) record;
			this.curRow = thisRow = numrec.getRow();
			thisColumn = numrec.getColumn();
			// 判断是否是日期格式,如果是日期则返回yyyy-MM-dd HH:mm:ss格式字符串
			double cellVlaue = numrec.getValue();
			if (DateUtil.isADateFormat(
					this.formatListener.getFormatIndex(numrec),
					this.formatListener.getFormatString(numrec))
					&& DateUtil.isValidExcelDate(cellVlaue)) {
				value = sdf.format(DateUtil.getJavaDate(cellVlaue));
			} else {
				value = this.formatListener.formatNumberDateCell(numrec).trim();
			}
			value = value.equals("") ? null : value;
			for (int i = this.rowList.size(); i < thisColumn; i++) {
				this.rowList.add(i, null);
			}
			this.rowList.add(thisColumn, value);
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
		if (thisRow != -1 && thisRow != this.lastRowNumber) {
			this.lastColumnNumber = -1;
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
			this.lastRowNumber = thisRow;
		}
		if (thisColumn > -1) {
			this.lastColumnNumber = thisColumn;
		}

		// 行结束时的操作
		if (record instanceof LastCellOfRowDummyRecord) {
			if (this.minColumns > 0) {
				// 列值重新置空
				if (this.lastColumnNumber == -1) {
					this.lastColumnNumber = 0;
				}
			}
			// 行结束时， 调用 optRows() 方法
			this.lastColumnNumber = -1;
			int tmpCols = this.rowList.size();
			if (this.curRow > this.titleRow && tmpCols < this.colSize) {
				for (int i = 0; i < this.colSize - tmpCols; i++) {
					this.rowList.add(this.rowList.size(), null);
				}
			}
			try {
				optRows(this.sheetIndex, this.curRow, this.rowList);
			} catch (SQLException e) {
				e.printStackTrace();
			}
			if (this.curRow == this.titleRow) {
				this.colSize = this.rowList.size();
			}
			this.rowList.clear();
		}
	}

	public int getTitleRow() {
		return this.titleRow;
	}

	public void setTitleRow(int titleRow) {
		this.titleRow = titleRow;
	}
}
