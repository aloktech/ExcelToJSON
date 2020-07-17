package com.imos.etj;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.util.LinkedHashMap;
import java.util.Stack;
import lombok.extern.log4j.Log4j2;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.json.JSONArray;
import org.json.JSONObject;

/**
 *
 * @author p
 */
@Log4j2
public class ExcelDataExtractor {

    private int rowCount;
    private int rowIndex;
    private int colIndex;
    private boolean forJSONArray;
    private Stack<String> stack;

    public JSONObject generateJSONDataFromExcelFile(String excelFilePath, String sheetName) {
        stack = new Stack<>();
        try (InputStream excelFile = new FileInputStream(excelFilePath);
                Workbook workbook = WorkbookFactory.create(excelFile)) {
            Sheet sheet = workbook.getSheet(sheetName);
            rowCount = sheet.getPhysicalNumberOfRows();
            rowIndex = 0;
            colIndex = 0;
            JSONObject jsonData = generateJSONObject(sheet);
            return jsonData;
        } catch (FileNotFoundException ex) {
            log.error(ex.getMessage());
        } catch (IOException ex) {
            log.error(ex.getMessage());
        }
        return new JSONObject();
    }

    JSONArray generateJSONArray(Sheet sheet) {
        return generateJSONArray(sheet, false);
    }

    JSONArray generateJSONArray(Sheet sheet, boolean jsonObjectArray) {
        JSONArray jsonData = new JSONArray();
        Row row;
        Object valueObj = "";
        if (jsonObjectArray) {
            while (forJSONArray) {
                jsonData.put(generateJSONObject(sheet));
                rowIndex++;
            }
            return jsonData;
        }
        rowLoop:
        while (rowIndex <= rowCount) {
            row = sheet.getRow(rowIndex);
            if (row != null) {
                int columnCount = sheet.getRow(rowIndex).getLastCellNum();
                colLoop:
                for (int columnIndex = colIndex; columnIndex < columnCount; columnIndex++) {
                    Cell cell = row.getCell(columnIndex, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                    if (cell == null) {
                        continue;
                    }
                    log.debug(rowIndex + ":" + columnIndex);
                    CellType cellType = cell.getCellType();
                    switch (cellType) {
                        case STRING: {
                            String value = cell.getStringCellValue().trim();
                            if (checkLineIsComment(value)) {
                                break colLoop;
                            }
                            if ("}".equals(value)) {
                                if (!stack.isEmpty() && "{".equals(stack.peek())) {
                                    stack.pop();
                                }

                                break rowLoop;
                            } else if ("},".equals(value)) {
                                if (!stack.isEmpty() && "{".equals(stack.peek())) {
                                    stack.pop();
                                }
                                break colLoop;
                            } else if ("},{".equals(value)) {
                                if (!stack.isEmpty() && "{".equals(stack.peek())) {
                                    stack.pop();
                                    stack.push("{");
                                }
                                break colLoop;
                            } else if ("]".equals(value)) {
                                if (!stack.isEmpty() && "[".equals(stack.peek())) {
                                    stack.pop();
                                }
                                break rowLoop;
                            } else if ("}]".equals(value)) {
                                if (!stack.isEmpty() && "{".equals(stack.peek())) {
                                    stack.pop();
                                }
                                if (!stack.isEmpty() && "[".equals(stack.peek())) {
                                    stack.pop();
                                }
                                break rowLoop;
                            }
                            if (stack.isEmpty()) {
                                break colLoop;
                            }
                            log.debug(value);
                            if ("{".equals(value)) {
                                stack.push(value);
                                rowIndex++;
                                jsonData.put(generateJSONObject(sheet));
                                break colLoop;
                            } else if ("[".equals(value)) {
                                stack.push(value);
                                rowIndex++;
                                jsonData.put(generateJSONArray(sheet));
                                break colLoop;
                            } else if ("[{".equals(value)) {
                                stack.push("[");
                                stack.push("{");
                                rowIndex++;
                                jsonData.put(generateJSONArray(sheet, true));
                                break colLoop;
                            } else {
                                valueObj = value;
                            }
                            break;
                        }
                        case BOOLEAN: {
                            valueObj = cell.getBooleanCellValue();
                            log.debug(valueObj);
                            break;
                        }
                        case NUMERIC: {
                            valueObj = cell.getNumericCellValue();
                            log.debug(valueObj);
                            break;
                        }
                    }
                    jsonData.put(valueObj);
                }
            }
            rowIndex++;
        }

        return jsonData;
    }

    private static boolean checkLineIsComment(String value) {
        return value.startsWith("//")
                || value.startsWith("##")
                || value.startsWith("#");
    }

    JSONObject generateJSONObject(Sheet sheet) {
        JSONObject jsonData = new JSONObject();
        setMapAsLinkedListMap(jsonData);
        Row row;
        String keyStr;
        Object valueObj = "";
        boolean firstRow = true;

        rowLoop:
        while (rowIndex <= rowCount) {
            row = sheet.getRow(rowIndex);
            if (row != null) {
                int columnCount = sheet.getRow(rowIndex).getLastCellNum();
                colLoop:
                for (int columnIndex = colIndex; columnIndex < columnCount; columnIndex++) {
                    Cell cell = row.getCell(columnIndex, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                    if (cell == null) {
                        continue;
                    }
                    log.debug(rowIndex + ":" + columnIndex);
                    CellType cellType = cell.getCellType();
                    switch (cellType) {
                        case STRING: {
                            keyStr = cell.getStringCellValue();
                            log.debug(keyStr);
                            if (checkLineIsComment(keyStr)) {
                                break colLoop;
                            }
                            if ("{".equals(keyStr)) {
                                stack.push(keyStr);
                                break colLoop;
                            } else if ("[".equals(keyStr)) {
                                stack.add(keyStr);
                                break colLoop;
                            } else if ("}".equals(keyStr)) {
                                columnIndex--;
                                if (!stack.isEmpty() && "{".equals(stack.peek())) {
                                    stack.pop();
                                }
                                break rowLoop;
                            } else if ("]".equals(keyStr)) {
                                columnIndex--;
                                if (!stack.isEmpty() && "[".equals(stack.peek())) {
                                    stack.pop();
                                }
                                break colLoop;
                            } else if ("}]".equals(keyStr)) {
                                columnIndex--;
                                if (!stack.isEmpty() && "{".equals(stack.peek())) {
                                    stack.pop();
                                }
                                if (!stack.isEmpty() && "[".equals(stack.peek())) {
                                    stack.pop();
                                }
                                forJSONArray = false;
                                break rowLoop;
                            }
                            if (stack.isEmpty()) {
                                break colLoop;
                            }
                            cell = row.getCell(columnIndex + 1, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                            if (cell == null) {
                                valueObj = "";
                            } else if (null != cell.getCellType()) {
                                log.debug(rowIndex + ":" + (columnIndex + 1));
                                switch (cell.getCellType()) {
                                    case STRING:
                                        String value = cell.getStringCellValue().trim();
                                        log.debug(value);
                                        if (checkLineIsComment(value)) {
                                            break colLoop;
                                        }
                                        if ("{".equals(value)) {
                                            stack.push(value);
                                            rowIndex++;
                                            jsonData.put(keyStr, generateJSONObject(sheet));
                                            break colLoop;
                                        } else if ("[".equals(value)) {
                                            forJSONArray = true;
                                            stack.push(value);
                                            rowIndex++;
                                            jsonData.put(keyStr, generateJSONArray(sheet));
                                            break colLoop;
                                        } else if ("[{".equals(value)) {
                                            forJSONArray = true;
                                            stack.push("[");
                                            stack.push("{");
                                            rowIndex++;
                                            jsonData.put(keyStr, generateJSONArray(sheet, true));
                                            rowIndex--;
                                            break colLoop;
                                        } else if ("}".equals(value)
                                                || "},".equals(value)
                                                || "} ,".equals(value)) {
                                            if (stack.isEmpty() && "{".equals(stack.peek())) {
                                                stack.pop();
                                            }
                                            break rowLoop;
                                        } else if ("]".equals(value)) {
                                            if (stack.isEmpty() && "[".equals(stack.peek())) {
                                                stack.pop();
                                            }
                                            break rowLoop;
                                        } else {
                                            valueObj = value;
                                        }
                                        break;
                                    case BOOLEAN:
                                        valueObj = cell.getBooleanCellValue();
                                        log.debug(valueObj);
                                        break;
                                    case NUMERIC:
                                        valueObj = cell.getNumericCellValue();
                                        log.debug(valueObj);
                                        break;
                                    default:
                                        valueObj = "";
                                        break;
                                }
                            }
                            log.debug(cell == null ? "BLANK" : cell.getCellType());
                            jsonData.put(keyStr, valueObj);
                            break colLoop;
                        }
                    }
                }
            }
            rowIndex++;
        }
        return jsonData;
    }

    private void setMapAsLinkedListMap(JSONObject json) {
        try {
            Field map = json.getClass().getDeclaredField("map");
            map.setAccessible(true);
            map.set(json, new LinkedHashMap<>());
            map.setAccessible(false);
        } catch (IllegalAccessException | IllegalArgumentException | NoSuchFieldException | SecurityException e) {
            log.error(e.getMessage());
        }
    }
}
