package com.huangxx.util.ExcelUtil;

import org.junit.Test;

import java.io.File;
import java.io.FileInputStream;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;
import java.util.Map;

/**
 * The <code>ImportExcelTests</code>
 */
public class ImportExcelTests {

    private static final SimpleDateFormat simpleDateFormat = new SimpleDateFormat("MM/dd/yy HH:mm");

    @Test
    public void testImport() throws Exception {
        File f = new File("D:\\T_COMB_ACC_VALUE.xlsx");
        FileInputStream inputStream = new FileInputStream(f);

        List<Map> importExcel = ExcelUtils.importExcel(Map.class, inputStream);

        String str = importExcel.get(1).get("CREATE_DATE").toString();
        Date date = simpleDateFormat.parse(str);
        for (Map m : importExcel) {
            System.out.println(m);
        }
    }

    @Test
    public void testImportBean() throws Exception {
        File f = new File("D:\\T_COMB_ACC_VALUE.xlsx");
        FileInputStream inputStream = new FileInputStream(f);
        List<TestModel> importExcel = ExcelUtils.importExcel(TestModel.class, inputStream);
    }
}
