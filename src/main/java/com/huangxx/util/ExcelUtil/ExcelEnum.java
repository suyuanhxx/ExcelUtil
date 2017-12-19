package com.huangxx.util.ExcelUtil;

import lombok.Getter;

/**
 * Created by huangxx on 2017/12/18.
 */
@Getter
public enum ExcelEnum {

    HSSF(".xls", "excel 2007"),
    XSSF(".xlsx", "excel 2010+");

    private String extension;
    private String desc;

    private ExcelEnum(String extension, String desc) {
        this.extension = extension;
        this.desc = desc;
    }

    public String getExtension() {
        return this.extension;
    }

    public String getDesc() {
        return this.desc;
    }
}
