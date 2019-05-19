/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.imos.etj;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.Iterator;
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
import org.json.JSONArray;
import org.json.JSONException;
import org.json.JSONObject;

/**
 *
 * @author p
 */
public class JSONGenerator {

    private static final Stack<String> PARENT_KEYS = new Stack<>();
    private static final Stack<JSONDataType> PARENT_TYPES = new Stack<>();
    private static final Map<String, JSONData> KEY_MAP = new LinkedHashMap<>();

    public void collectFromExcelFile(String excelFilePath, String sheetName) {
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
                                    break;
                                }
                                cell = row.getCell(j + 1);
                                if (cell == null) {
                                    JSONDataType parentType = setParentType();
                                    if (null != parentType) {
                                        switch (parentType) {
                                            case ARRAY_VALUE:
                                                KEY_MAP.put(keyValue, new JSONData(
                                                        JSONDataType.VALUE, JSONValueType.STRING,
                                                        keyValue, keyValue, parentKey, setParentType()));
                                                break;
                                            case ARRAY_OBJECT:
                                                KEY_MAP.put(keyValue, new JSONData(
                                                        JSONDataType.VALUE, JSONValueType.STRING,
                                                        "Dummy", keyValue, parentKey, setParentType()));
                                                break;
                                            case OBJECT:
                                                KEY_MAP.put(keyValue, new JSONData(
                                                        JSONDataType.VALUE, JSONValueType.STRING,
                                                        "Dummy", keyValue, parentKey, setParentType()));
                                                break;
                                        }
                                    } else {
                                        KEY_MAP.put(keyValue, new JSONData(
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
                                KEY_MAP.put(String.valueOf(booleanValue), new JSONData(
                                        JSONDataType.ARRAY_VALUE, JSONValueType.BOOLEAN,
                                        booleanValue, keyValue, parentKey, setParentType()));
                            } else {
                                KEY_MAP.put(keyValue, new JSONData(
                                        JSONDataType.VALUE, JSONValueType.BOOLEAN,
                                        booleanValue, keyValue, parentKey, setParentType()));
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

    public JSONTreeNode setJSONValue() {
        JSONTreeNode root = new JSONTreeNode(null, null);
        KEY_MAP.forEach((key, value) -> {
            if (value.getParentKey() == null) {
                root.getChildren().add(new JSONTreeNode(value, null));
            } else {
                searchTreeAndSetValue(root, value);
            }
        });
        return root;
    }

    private void collectStringValue(Cell cell, String keyValue, String parentKey) {
        String stringValue = cell.getStringCellValue();
        if (null == stringValue) {
            KEY_MAP.put(keyValue, new JSONData(
                    JSONDataType.VALUE, JSONValueType.STRING,
                    stringValue, keyValue, parentKey, setParentType()));
        } else {
            switch (stringValue) {
                case "{":
                    PARENT_KEYS.push(keyValue);
                    KEY_MAP.put(keyValue, new JSONData(
                            JSONDataType.OBJECT, JSONValueType.OBJECT,
                            null, keyValue, parentKey, setParentType()));
                    PARENT_TYPES.push(JSONDataType.OBJECT);
                    break;
                case "[{":
                    PARENT_KEYS.push(keyValue);
                    KEY_MAP.put(keyValue, new JSONData(
                            JSONDataType.ARRAY_OBJECT, JSONValueType.OBJECT,
                            null, keyValue, parentKey, setParentType()));
                    PARENT_TYPES.push(JSONDataType.ARRAY_OBJECT);
                    break;
                case "[":
                    PARENT_KEYS.push(keyValue);
                    KEY_MAP.put(keyValue, new JSONData(
                            JSONDataType.ARRAY_VALUE, JSONValueType.STRING,
                            null, keyValue, parentKey, setParentType()));
                    PARENT_TYPES.push(JSONDataType.ARRAY_VALUE);
                    break;
                default:
                    KEY_MAP.put(keyValue, new JSONData(
                            JSONDataType.VALUE, JSONValueType.STRING,
                            stringValue, keyValue, parentKey, setParentType()));
                    break;
            }
        }
    }

    private void collectNumericData(String strData, String keyValue, String parentKey, Cell cell) {
        int intValue;
        long longValue;
        double doubleValue;
        try {
            intValue = Integer.parseInt(strData);
            KEY_MAP.put(keyValue == null ? strData : keyValue, new JSONData(
                    JSONDataType.VALUE, JSONValueType.INTEGER,
                    intValue, keyValue, parentKey, setParentType()));
        } catch (NumberFormatException e1) {
            try {
                longValue = Long.parseLong(strData);
                KEY_MAP.put(keyValue == null ? strData : keyValue, new JSONData(
                        JSONDataType.VALUE, JSONValueType.LONG,
                        longValue, keyValue, parentKey, setParentType()));
                if (HSSFDateUtil.isCellDateFormatted(cell)) {
                    KEY_MAP.put(keyValue, new JSONData(
                            JSONDataType.VALUE, JSONValueType.DATE,
                            longValue, keyValue, parentKey, setParentType()));
                }
            } catch (NumberFormatException e2) {
                try {
                    doubleValue = Double.parseDouble(strData);
                    KEY_MAP.put(keyValue == null ? strData : keyValue, new JSONData(
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

    private JSONDataType setParentType() {
        return PARENT_TYPES.isEmpty() ? null : PARENT_TYPES.peek();
    }

    public void buildJSONObject(JSONTreeNode root, JSONObject jsonResult, JSONArray arrayResult) throws JSONException {
        for (JSONTreeNode node : root.getChildren()) {
            if (null != node.getValue().getDataType()) {
                switch (node.getValue().getDataType()) {
                    case VALUE:
                        Object obj = node.getValue().getData();
                        JSONDataType dataType = node.getValue().getParentDataType();
                        if (dataType == null) {
                            jsonResult.put(node.getValue().getKey(), setValue(obj));
                            continue;
                        }
                        switch (node.getValue().getParentDataType()) {
                            case OBJECT:
                                jsonResult.put(node.getValue().getKey(), setValue(obj));
                                break;
                            case ARRAY_VALUE:
                                arrayResult.put(node.getValue().getData());
                                break;
                            case ARRAY_OBJECT:
                                if (!PARENT_TYPES.isEmpty()) {
                                    JSONDataType parentDataType = PARENT_TYPES.peek();
                                    if (parentDataType == JSONDataType.ARRAY_OBJECT) {
                                        JSONObject json = new JSONObject();
                                        json.put(node.getValue().getKey(), node.getValue().getData());
                                        arrayResult.put(json);
                                    } else {
                                        arrayResult.put(node.getValue().getData());
                                    }
                                }
                                break;
                            default:
                                jsonResult.put(node.getValue().getKey(), setValue(obj));
                                break;
                        }
                        break;
                    case OBJECT:
                        JSONObject json = new JSONObject();
                        PARENT_TYPES.add(JSONDataType.OBJECT);
                        jsonResult.put(node.getValue().getKey(), json);
                        if (!node.getChildren().isEmpty()) {
                            buildJSONObject(node, json, arrayResult);
                        }
                        break;
                    case ARRAY_OBJECT:
                        JSONArray jsonArrayObject = new JSONArray();
                        JSONObject jsonObject = new JSONObject();
                        PARENT_TYPES.add(JSONDataType.ARRAY_OBJECT);
                        jsonResult.put(node.getValue().getKey(), jsonArrayObject);
                        if (!node.getChildren().isEmpty()) {
                            buildJSONObject(node, jsonObject, jsonArrayObject);
                        }
                        break;
                    case ARRAY_VALUE:
                        JSONArray jsonArrayValue = new JSONArray();
                        PARENT_TYPES.add(JSONDataType.ARRAY_VALUE);
                        jsonResult.put(node.getValue().getKey(), jsonArrayValue);
                        if (!node.getChildren().isEmpty()) {
                            buildJSONObject(node, jsonResult, jsonArrayValue);
                        }
                        break;
                    default:
                        break;
                }
            }
        }
    }

    private Object setValue(Object obj) {
        return obj == null ? "Dummy" : obj;
    }

    private void searchTreeAndSetValue(JSONTreeNode root, JSONData value) {
        Iterator<JSONTreeNode> itr = root.getChildren().iterator();
        while (itr.hasNext()) {
            JSONTreeNode treeNode = itr.next();
            if (treeNode.getValue().getKey() != null
                    && value.getParentKey() != null
                    && treeNode.getValue().getKey().equals(value.getParentKey())) {
                treeNode.getChildren().add(new JSONTreeNode(value, treeNode.getValue()));
            } else if (!treeNode.getChildren().isEmpty()) {
                searchTreeAndSetValue(treeNode, value);
            }
        }
    }
}
