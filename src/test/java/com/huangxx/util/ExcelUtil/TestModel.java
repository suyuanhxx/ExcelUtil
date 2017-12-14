package com.huangxx.util.ExcelUtil;

import lombok.Data;

import java.util.Date;

/**
 * The <code>Model</code>
 */
@Data
public class TestModel {
    @ExcelCell(index = 0)
    private String a;
    @ExcelCell(index = 1)
    private String b;
    @ExcelCell(index = 2)
    private String c;
    @ExcelCell(index = 3)
    private Date d;

}
