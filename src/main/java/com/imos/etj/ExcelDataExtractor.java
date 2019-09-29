package com.imos.etj;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.Stack;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

/**
 *
 * @author p
 */
public class ExcelDataExtractor {

    private static final Stack<String> PARENT_KEYS = new Stack<>();
    private static final Stack<JSONDataType> PARENT_TYPES = new Stack<>();

    private final Map<String, JSONData> KEY_MAP = new LinkedHashMap<>();

    int booleanRowIndex = 0;

    public void collectDataFromExcelFile(String excelFilePath, String sheetName) {
        try (InputStream excelFile = new FileInputStream(excelFilePath);
                Workbook workbook = WorkbookFactory.create(excelFile)) {
            Sheet sheet = workbook.getSheet(sheetName);
            int rowCount = sheet.getLastRowNum();
            DataFormatter fmt = new DataFormatter();
            for (int i = 0; i <= rowCount; i++) {
                Row row = sheet.getRow(i);
                if (row == null) {
                    continue;
                }
                int columnCount = row.getLastCellNum();
                String parentKey, keyValue = null;
                boolean booleanValue;
                for (int j = 0; j < columnCount; j++) {
                    Cell cell = row.getCell(j, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                    if (cell == null) {
                        continue;
                    }
                    CellType cellType = cell.getCellType();
                    parentKey = PARENT_KEYS.isEmpty() ? null : PARENT_KEYS.peek();
                    switch (cellType) {
                        case STRING:
                            if (keyValue == null) {
                                keyValue = cell.getStringCellValue();
                                if ("}".equals(keyValue) || "}]".equals(keyValue) || "]".equals(keyValue)) {
                                    if (!PARENT_KEYS.isEmpty()) {
                                        PARENT_KEYS.pop();
                                    }
                                    if (!PARENT_TYPES.isEmpty()) {
                                        PARENT_TYPES.pop();
                                    }
                                    booleanRowIndex = 0;
                                    break;
                                }
                                cell = row.getCell(j + 1);
                                if (cell == null) {
                                    JSONDataType parentType = setParentType();
                                    if (null != parentType) {
                                        switch (parentType) {
                                            case ARRAY_VALUE:
                                                KEY_MAP.put(getKeyValue(KEY_MAP, keyValue), new JSONData(
                                                        JSONDataType.VALUE, JSONValueType.ARRAY,
                                                        keyValue, keyValue, parentKey, setParentType()));
                                                break;
                                            case ARRAY_OBJECT:
                                                KEY_MAP.put(getKeyValue(KEY_MAP, keyValue), new JSONData(
                                                        JSONDataType.VALUE, JSONValueType.ARRAY,
                                                        "Dummy", keyValue, parentKey, setParentType()));
                                                break;
                                            case OBJECT:
                                                KEY_MAP.put(getKeyValue(KEY_MAP, keyValue), new JSONData(
                                                        JSONDataType.VALUE, JSONValueType.STRING,
                                                        "Dummy", keyValue, parentKey, setParentType()));
                                                break;
                                        }
                                    } else {
                                        KEY_MAP.put(getKeyValue(KEY_MAP, keyValue), new JSONData(
                                                JSONDataType.VALUE, JSONValueType.STRING,
                                                "Dummy", keyValue, parentKey, setParentType()));
                                    }
                                } else {
                                    cellType = cell.getCellType();
                                    if (cellType == CellType.STRING) {
                                        collectStringValue(cell, keyValue, parentKey);
                                    }
                                }
                            }
                            break;
                        case BOOLEAN:
                            booleanValue = cell.getBooleanCellValue();
                            if (keyValue == null) {
                                keyValue = parentKey;
                                KEY_MAP.put(getKeyValue(KEY_MAP, keyValue, booleanRowIndex), new JSONData(
                                        JSONDataType.VALUE, JSONValueType.BOOLEAN,
                                        booleanValue, keyValue, parentKey, setParentType()));
                                booleanRowIndex++;
                            } else {
                                KEY_MAP.put(getKeyValue(KEY_MAP, keyValue, booleanRowIndex), new JSONData(
                                        JSONDataType.VALUE, JSONValueType.BOOLEAN,
                                        booleanValue, keyValue, parentKey, setParentType()));
                                booleanRowIndex++;
                            }
                            break;
                        case NUMERIC:
                            String strData = fmt.formatCellValue(cell);
                            collectNumericData(strData, keyValue, parentKey, cell);
                            break;
                        default:
                            break;
                    }
                }
            }
        } catch (FileNotFoundException ex) {
            Logger.getLogger(ExcelToJSON.class.getName()).log(Level.SEVERE, null, ex);
        } catch (IOException ex) {
            Logger.getLogger(ExcelToJSON.class.getName()).log(Level.SEVERE, null, ex);
        }
    }

    private static String getKeyValue(Map<String, JSONData> KEY_MAP, String keyValue) {
        String key = keyValue;
        while (KEY_MAP.containsKey(key)) {
            key = key + "a";
        }
        return key;
    }

    private static String getKeyValue(Map<String, JSONData> KEY_MAP, Number keyValue, String defaultValue) {
        String keyValueStr = keyValue == null ? defaultValue : "" + keyValue;
        String key = keyValueStr;
        while (KEY_MAP.containsKey(key)) {
            key = key + "a";
        }
        return key;
    }

    private static String getKeyValue(Map<String, JSONData> KEY_MAP, String keyValue, int index) {
        String key = keyValue + ":" + index;
        while (KEY_MAP.containsKey(key)) {
            key = key + "a";
        }
        return key;
    }

    private JSONDataType setParentType() {
        return PARENT_TYPES.isEmpty() ? null : PARENT_TYPES.peek();
    }

    private void collectNumericData(String strData, String keyValue, String parentKey, Cell cell) {
        int intValue;
        long longValue;
        double doubleValue;
        try {
            intValue = Integer.parseInt(strData);
            KEY_MAP.put(getKeyValue(KEY_MAP, intValue, strData), new JSONData(
                    JSONDataType.VALUE, JSONValueType.INTEGER,
                    intValue, keyValue, parentKey, setParentType()));
        } catch (NumberFormatException e1) {
            try {
                longValue = Long.parseLong(strData);
                KEY_MAP.put(getKeyValue(KEY_MAP, longValue, strData), new JSONData(
                        JSONDataType.VALUE, JSONValueType.LONG,
                        longValue, keyValue, parentKey, setParentType()));
                if (HSSFDateUtil.isCellDateFormatted(cell)) {
                    KEY_MAP.put(getKeyValue(KEY_MAP, keyValue), new JSONData(
                            JSONDataType.VALUE, JSONValueType.DATE,
                            longValue, keyValue, parentKey, setParentType()));
                }
            } catch (NumberFormatException e2) {
                try {
                    doubleValue = Double.parseDouble(strData);
                    KEY_MAP.put(getKeyValue(KEY_MAP, doubleValue, strData), new JSONData(
                            JSONDataType.VALUE, JSONValueType.DOUBLE,
                            doubleValue, keyValue, parentKey, setParentType()));
                } catch (NumberFormatException e3) {
                    KEY_MAP.put(keyValue, new JSONData(
                            JSONDataType.VALUE, JSONValueType.NONE,
                            null, keyValue, parentKey, setParentType()));
                }
            }
        }
    }

    private void collectStringValue(Cell cell, String keyValue, String parentKey) {
        String stringValue = cell.getStringCellValue();
        if (null == stringValue) {
            KEY_MAP.put(getKeyValue(KEY_MAP, keyValue), new JSONData(
                    JSONDataType.VALUE, JSONValueType.STRING,
                    stringValue, keyValue, parentKey, setParentType()));
        } else {
            switch (stringValue) {
                case "{":
                    PARENT_KEYS.push(keyValue);
                    KEY_MAP.put(getKeyValue(KEY_MAP, keyValue), new JSONData(
                            JSONDataType.OBJECT, JSONValueType.OBJECT,
                            null, keyValue, parentKey, setParentType()));
                    PARENT_TYPES.push(JSONDataType.OBJECT);
                    break;
                case "[{":
                    PARENT_KEYS.push(keyValue);
                    KEY_MAP.put(getKeyValue(KEY_MAP, keyValue), new JSONData(
                            JSONDataType.ARRAY_OBJECT, JSONValueType.ARRAY,
                            null, keyValue, parentKey, setParentType()));
                    PARENT_TYPES.push(JSONDataType.ARRAY_OBJECT);
                    break;
                case "[":
                    booleanRowIndex = 0;
                    PARENT_KEYS.push(keyValue);
                    KEY_MAP.put(getKeyValue(KEY_MAP, keyValue), new JSONData(
                            JSONDataType.ARRAY_VALUE, JSONValueType.ARRAY,
                            null, keyValue, parentKey, setParentType()));
                    PARENT_TYPES.push(JSONDataType.ARRAY_VALUE);
                    break;
                default:
                    KEY_MAP.put(getKeyValue(KEY_MAP, keyValue), new JSONData(
                            JSONDataType.VALUE, JSONValueType.STRING,
                            stringValue, keyValue, parentKey, setParentType()));
                    break;
            }
        }
    }

    public Map<String, JSONData> getJSONKeyValueMap() {
        return KEY_MAP;
    }

}
