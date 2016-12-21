package module.cartype;

import java.util.List;
import java.util.Map;

import org.nutz.dao.Dao;

import dao.DaoUtil;
import excel.ExcelReaderUtil;
import excel.read.IRowReader;
import excel.utils.ExcelColumn;
import excel.utils.GlobalContext;

public class CarTypeExe {
    public static void main(String[] args) {
        run();
    }
    int index = 10;
    private static void run() {
        Dao dao = DaoUtil.getDao();
        long start = System.currentTimeMillis();
        try {
            IRowReader<Object> reader = new IRowReader<Object>() {
                @Override
                public Map<String, Object> getRows(int sheetIndex, int curRow, List<Object> rowList) {
//                    System.out.print(curRow + " ");
//                    for (int i = 0; i < rowList.size(); i++) {
//                        System.out.print(rowList.get(i) + ",");
//                    }
//                    System.out.println();
                    //
                    CarType carType = new CarType();
                    carType.setLevel_id(rowList.get(ExcelColumn.excelColStrToNum("A")).toString());
                    carType.setYpc_id(0);
                    carType.setInitials(rowList.get(ExcelColumn.excelColStrToNum("GG")).toString());
                    carType.setBrand(rowList.get(ExcelColumn.excelColStrToNum("GH")).toString());
                    carType.setFactory(rowList.get(ExcelColumn.excelColStrToNum("GI")).toString());
                    carType.setCars(rowList.get(ExcelColumn.excelColStrToNum("GJ")).toString());
                    carType.setYear(rowList.get(ExcelColumn.excelColStrToNum("GL")).toString());
                    carType.setDisplacement(rowList.get(ExcelColumn.excelColStrToNum("GN")).toString());
                    carType.setSale_name(rowList.get(ExcelColumn.excelColStrToNum("GG")).toString());
                    carType.setCar_type(rowList.get(ExcelColumn.excelColStrToNum("GK")).toString());
                    carType.setYear_model(rowList.get(ExcelColumn.excelColStrToNum("GM")).toString());
                    carType.setChassis(null);
                    carType.setEngine(rowList.get(ExcelColumn.excelColStrToNum("GR")).toString());
                    carType.setIntake_type(rowList.get(ExcelColumn.excelColStrToNum("G0")).toString());
                    carType.setGearbox_type(rowList.get(ExcelColumn.excelColStrToNum("GP")).toString());
                    carType.setGearbox_remark(rowList.get(ExcelColumn.excelColStrToNum("GQ")).toString());
                    carType.setBrand_logo_small(rowList.get(ExcelColumn.excelColStrToNum("GG")).toString());
                    return null;
                }
            };
            new ExcelReaderUtil().readExcel(reader, GlobalContext.EXCEL07_EXTENSION, "E:/易配诚改版.xlsx", null, 1, 0, 1, "yyyyMMdd");
            System.out.println(System.currentTimeMillis() - start);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
