package com.huangxx.util.ExcelUtil;

import org.junit.Test;

import java.io.File;
import java.io.FileInputStream;
import java.util.List;

public class ExportExcelTests {

    @Test
    public void testExport() throws Exception {
        File f = new File("D:\\T_COMB_ACC_VALUE.xlsx");
        FileInputStream inputStream = new FileInputStream(f);
        List<TestModel> importExcel = ExcelUtils.importExcel(TestModel.class, inputStream);
    }

    @Test
    public void testExportBean() throws Exception {
        File f = new File("D:\\T_COMB_ACC_VALUE.xlsx");
        FileInputStream inputStream = new FileInputStream(f);
        List<TestModel> importExcel = ExcelUtils.importExcel(TestModel.class, inputStream);
    }

}
