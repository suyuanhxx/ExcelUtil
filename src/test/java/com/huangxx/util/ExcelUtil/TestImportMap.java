package com.huangxx.util.ExcelUtil;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.InputStream;
import java.util.Collection;
import java.util.Map;

/**
 * The <code>TestImportMap</code>
 *
 */
public class TestImportMap {

    public static void main(String[] args) throws FileNotFoundException {
        File f = new File("E:\\July.xlsx");
        InputStream inputStream = new FileInputStream(f);

        Collection<Map> importExcel = ExcelUtils.importExcel(Map.class, inputStream, "yyyy/MM/dd HH:mm:ss");

        for (Map m : importExcel) {
            System.out.println(m);
        }
    }
}
