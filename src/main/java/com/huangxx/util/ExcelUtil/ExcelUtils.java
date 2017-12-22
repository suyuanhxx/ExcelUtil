package com.huangxx.util.ExcelUtil;

import org.apache.commons.collections.CollectionUtils;
import org.apache.commons.collections4.MapUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.log4j.Logger;
import org.apache.log4j.LogManager;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.text.MessageFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Locale;
import java.util.Map;

/**
 * The <code>ExcelUtil</code> 与 {@link ExcelCell}搭配使用
 * Created by huangxx on 2017/11/22.
 */
public class ExcelUtils {

    private static final Logger log = LogManager.getLogger(ExcelUtils.class);

    private static final SimpleDateFormat simpleDateFormat = new SimpleDateFormat("MM/dd/yy HH:mm");

    /**
     * 用来验证excel与Vo中的类型是否一致 <br>
     * Map<栏位类型,只能是哪些Cell类型>
     */
    private static DataFormatter dataFormatter = new DataFormatter(Locale.CHINA);

    public static Workbook generateWorkBook(ExcelEnum excelEnum, String sheetName, List<String> headers, List<List<String>> dataRows) {
        validateParams(headers, dataRows);
        Workbook workbook = null;
        try {
            if (excelEnum == ExcelEnum.HSSF) {
                workbook = new HSSFWorkbook();
            } else if (excelEnum == ExcelEnum.XSSF) {
                workbook = new XSSFWorkbook();
            }
            if (workbook == null) {
                return null;
            }

            Sheet sheet = workbook.createSheet(sheetName);
            fillSheet(sheet, headers, dataRows);
            return workbook;
        } catch (Exception var6) {
            throw new RuntimeException("failed to generate WorkBook " + var6.getMessage(), var6);
        }
    }

    public static OutputStream write2OutputSteam(ExcelEnum excelEnum, String sheetName, List<String> headers, List<List<String>> dataRows) {
        Workbook workbook = generateWorkBook(excelEnum, sheetName, headers, dataRows);
        if (workbook == null) {
            return null;
        }
        OutputStream bytes = new ByteArrayOutputStream();
        try {
            workbook.write(bytes);
            return bytes;
        } catch (Exception e) {
            log.error("write2OutputSteam error!", e);
            return null;
        }
    }


    /**
     * @param excelEnum 文件类型
     * @param headers   title
     * @param dataRows  数据
     * @param desPath   desPath文件路径
     */
    public static void exportExcel(ExcelEnum excelEnum, List<String> headers, List<List<String>> dataRows, String desPath) {
        validateParams(headers, dataRows, desPath);
        Workbook workbook = null;
        FileOutputStream outputStream = null;

        try {
            if (excelEnum == ExcelEnum.HSSF) {
                workbook = new HSSFWorkbook();
            } else if (excelEnum == ExcelEnum.XSSF) {
                workbook = new XSSFWorkbook();
            }
            if (workbook == null) {
                return;
            }

            Sheet sheet = workbook.createSheet();
            fillSheet(sheet, headers, dataRows);
            outputStream = new FileOutputStream(desPath);
            workbook.write(outputStream);
        } catch (Exception e) {
            throw new RuntimeException("failed to generate WorkBook " + e.getMessage(), e);
        } finally {
            try {
                IOUtils.closeQuietly(outputStream);
                assert workbook != null;
                workbook.close();
            } catch (Exception e) {
                log.error("workbook close error!", e);
            }

        }

    }

    /**
     * 把Excel的数据封装成voList
     *
     * @param clazz       vo的Class
     * @param inputStream excel输入流
     * @return voList
     * @throws RuntimeException
     */
    public static <T> List<T> importExcel(Class<T> clazz, InputStream inputStream) {
        Workbook workbook = null;
        try {
            workbook = WorkbookFactory.create(inputStream);
        } catch (Exception e) {
            log.error("importExcel create workbook error", e);
        }
        if (workbook == null) {
            return null;
        }

        List<T> list = new ArrayList<>();
        Sheet sheet = workbook.getSheetAt(0);
        Iterator<Row> rowIterator = sheet.rowIterator();

        if (clazz == Map.class) {
            Map<String, Integer> titleMap = new HashMap<>();
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                if (row.getRowNum() == 0) {
                    titleMap = getHeaderTitleMap(row);
                    continue;
                }
                Map<String, Object> map = initMapValue(row, titleMap);
                if (MapUtils.isEmpty(map)) {
                    continue;
                }
                list.add((T) map);
            }
        } else {
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                boolean allRowIsNull = isEmptyLine(row);
                if (allRowIsNull) {
                    continue;
                }
                T t = initObjectValue(clazz, row);
                if (t == null) {
                    continue;
                }
                list.add(t);
            }
        }
        return list;
    }

    private static Map<String, Integer> getHeaderTitleMap(Row row) {
        Map<String, Integer> titleMap = new HashMap<>();
        Integer index = 0;
        Iterator<Cell> cellIterator = row.cellIterator();
        while (cellIterator.hasNext()) {
            String value = cellIterator.next().getStringCellValue();
            titleMap.put(value, index);
            index++;
        }
        return titleMap;
    }

    /**
     * 判断整行是否为空
     *
     * @param row row
     * @return
     */
    private static boolean isEmptyLine(Row row) {
        boolean allRowIsNull = true;
        Iterator<Cell> cellIterator = row.cellIterator();
        while (cellIterator.hasNext()) {
            Object cellValue = dataFormatter.formatCellValue(cellIterator.next());
            if (cellValue != null) {
                allRowIsNull = false;
                break;
            }
        }
        return allRowIsNull;
    }

    private static Map<String, Object> initMapValue(Row row, Map<String, Integer> titleMap) {
        Map<String, Object> map = new HashMap<>();
        for (String k : titleMap.keySet()) {
            Integer index = titleMap.get(k);
            Cell cell = row.getCell(index);
            String value = dataFormatter.formatCellValue(cell);
            if (StringUtils.isBlank(value)) {
                continue;
            }
            map.put(k, value);
        }
        return map;
    }

    private static <T> T initObjectValue(Class<T> clazz, Row row) {
        try {
            T t = clazz.newInstance();
            Field[] fields = clazz.getDeclaredFields();
            for (Field field : fields) {
                field.setAccessible(true);
                ExcelCell annotation = field.getAnnotation(ExcelCell.class);
                int cellIndex = annotation.index();
                initValue(t, field, row.getCell(cellIndex));
            }
            return t;
        } catch (Exception e) {
            log.error("can not instance class:" + clazz.getSimpleName(), e);
            throw new RuntimeException(MessageFormat.format("can not instance class:{0}",
                    clazz.getSimpleName()), e);
        }
    }

    private static <T> void initValue(T t, Field field, Cell cell) {
        if (cell == null) {
            return;
        }
        try {
            if (field.getType() == String.class) {
                field.set(t, cell.getStringCellValue());
                return;
            }
            if (field.getType() == Double.class) {
                field.set(t, cell.getNumericCellValue());
                return;
            }
            if (field.getType() == Date.class) {
                field.set(t, cell.getDateCellValue());
                return;
            }
            field.set(t, cell.getStringCellValue());
        } catch (Exception e) {
            log.error("initValue error", e);
        }
    }

    private static void validateParams(List<String> headers, List<List<String>> dataRows, String desPath) {
        if (CollectionUtils.isEmpty(dataRows)) {
            throw new RuntimeException("headers is empty ");
        } else {
            validateCommonParams(dataRows, desPath);
        }
    }

    private static void validateParams(List<String> headers, List<List<String>> dataRows) {
        if (CollectionUtils.isEmpty(headers)) {
            throw new RuntimeException("headers is empty ");
        } else if (CollectionUtils.isEmpty(dataRows)) {
            throw new RuntimeException("dataRows is empty ");
        }
    }

    private static void validateCommonParams(List<List<String>> dataRows, String desPath) {
        if (CollectionUtils.isEmpty(dataRows)) {
            throw new RuntimeException("dataRows is empty");
        } else if (StringUtils.isBlank(desPath)) {
            throw new RuntimeException("desPath is empty");
        } else {
            String dir = desPath.substring(0, desPath.lastIndexOf("/"));
            File file = new File(dir);
            if (!file.exists() && !file.mkdirs()) {
                throw new RuntimeException("failed to create dir " + dir);
            }
        }
    }

    private static void fillSheet(Sheet sheet, List<String> headers, List<List<String>> dataRows) {
        createRow(sheet, 0, headers);
        fillSheet(sheet, dataRows);
    }

    private static void fillSheet(Sheet sheet, List<List<String>> dataRows) {
        for (int i = 1; i <= dataRows.size(); ++i) {
            createRow(sheet, i, dataRows.get(i - 1));
        }
    }

    private static void createRow(Sheet sheet, int rowNum, List<String> rowData) {
        Row row = sheet.createRow(rowNum);

        for (int i = 0; i < rowData.size(); ++i) {
            Cell cell = row.createCell(i);
            cell.setCellType(1);
            cell.setCellValue((String) rowData.get(i));
        }
    }

}
