package excel.read;

import java.util.List;
import java.util.Map;

public class RowReader implements IRowReader<Object> {

	/*
	 * 业务逻辑实现方法
	 * 
	 * @see com.eprosun.util.excel.IRowReader#getRows(int, int, java.util.List)
	 */
	@Override
	public Map<String, Object> getRows(int sheetIndex, int curRow,
			List<Object> rowList) {

		System.out.print(curRow + " ");
		for (int i = 0; i < rowList.size(); i++) {
			System.out.print(rowList.get(i) + ",");
		}
		System.out.println();
		return null;
	}

}
