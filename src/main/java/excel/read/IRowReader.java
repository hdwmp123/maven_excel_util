package excel.read;

import java.util.List;
import java.util.Map;

public interface IRowReader<T> {

	/**
	 * 业务逻辑实现方法
	 * 
	 * @param sheetIndex
	 * @param curRow
	 * @param rowList
	 */
	public Map<String, Object> getRows(int sheetIndex, int curRow,
			List<T> rowList);
}
