package excel.test;

import java.util.List;
import java.util.Map;

import excel.ExcelReaderUtil;
import excel.read.IRowReader;
import excel.utils.GlobalContext;

public class Test {
	public static void main(String[] args) {
		// test();
		test1();
	}

	private static void test1() {
		System.out.println(String.format("%.2f%s", ((double) 5 * 100 / 7),"%"));
	}

	private static void test() {
		long start = System.currentTimeMillis();
		try {
			new ExcelReaderUtil()
					.readExcel(
							new IRowReader<Object>() {
								@Override
								public Map<String, Object> getRows(
										int sheetIndex, int curRow,
										List<Object> rowList) {
									System.out.print(curRow + " ");
									for (int i = 0; i < rowList.size(); i++) {
										System.out.print(rowList.get(i) + ",");
									}
									System.out.println();
									return null;
								}
							},
							GlobalContext.EXCEL03_EXTENSION,
							"C:/Users/cwhl-0045/Desktop/Y医疗项目/数据导入/05月29日第二次数项目数 十八省稿/福建 20100529.xls",
							null, 1, 0, 1, "yyyyMMdd");
			System.out.println(System.currentTimeMillis() - start);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

}
