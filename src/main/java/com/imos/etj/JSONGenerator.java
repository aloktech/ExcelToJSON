package com.imos.etj;

import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.Stack;
import org.json.JSONArray;
import org.json.JSONException;
import org.json.JSONObject;

/**
 *
 * @author p
 */
public class JSONGenerator {

    private static final Stack<Object> PARENT_OBJECT = new Stack<>();
    private static final Stack<String> PARENT_KEYS = new Stack<>();
    private static final Stack<JSONDataType> PARENT_TYPES = new Stack<>();

    private static Map<String, JSONData> KEY_MAP = new LinkedHashMap<>();

    public JSONTreeNode setJSONValue(Map<String, JSONData> keyMap) {
        JSONTreeNode root = new JSONTreeNode(null, null);
        KEY_MAP = keyMap;
        KEY_MAP.forEach((key, value) -> {
            if (value.getParentKey() == null) {
                root.getChildren().add(new JSONTreeNode(value, null));
            } else {
                searchTreeAndSetValue(root, value);
            }
        });
        return root;
    }

    public void buildJSONObject(JSONTreeNode root, JSONObject jsonResult, JSONArray arrayResult) throws JSONException {
//        JSONDataType parentType = JSONDataType.ARRAY_OBJECT;
        JSONArray parentArray = new JSONArray();
        JSONObject parentJson = new JSONObject();
        for (JSONTreeNode node : root.getChildren()) {
            JSONData jsonData = node.getValue();
            if (null != jsonData.getDataType()) {
                switch (jsonData.getDataType()) {
                    case VALUE:
                        Object obj = jsonData.getData();
                        JSONDataType dataType = jsonData.getParentDataType();
                        if (dataType == null) {
                            jsonResult.put(jsonData.getKey(), setValue(obj));
                            continue;
                        }
                        switch (dataType) {
                            case OBJECT:
                                jsonResult.put(jsonData.getKey(), setValue(obj));
                                break;
                            case ARRAY_VALUE:
                                arrayResult.put(jsonData.getData());
                                break;
                            case ARRAY_OBJECT:
                                jsonResult.put(jsonData.getKey(), jsonData.getData());
                                break;
                            default:
                                jsonResult.put(jsonData.getKey(), setValue(obj));
                                break;
                        }
                        break;
                    case OBJECT:
                        JSONObject json = new JSONObject();
                        PARENT_TYPES.add(JSONDataType.OBJECT);
                        jsonResult.put(jsonData.getKey(), json);
                        if (!node.getChildren().isEmpty()) {
                            buildJSONObject(node, json, arrayResult);
                        }
                        break;
                    case ARRAY_OBJECT:
                        JSONArray jsonArrayObject = new JSONArray();
                        JSONObject jsonObject = new JSONObject();
                        jsonArrayObject.put(jsonObject);
                        PARENT_TYPES.add(JSONDataType.ARRAY_OBJECT);
                        jsonResult.put(jsonData.getKey(), jsonArrayObject);
                        if (!node.getChildren().isEmpty()) {
                            buildJSONObject(node, jsonObject, jsonArrayObject);
                        }
                        break;
                    case ARRAY_VALUE:
                        JSONArray jsonArrayValue = new JSONArray();
                        PARENT_TYPES.add(JSONDataType.ARRAY_VALUE);
                        jsonResult.put(jsonData.getKey(), jsonArrayValue);
                        if (!node.getChildren().isEmpty()) {
                            buildJSONObject(node, jsonResult, jsonArrayValue);
                        }
                        break;
                    default:
                        break;
                }
            }
        }
        if (!parentJson.isEmpty() && !PARENT_KEYS.isEmpty()) {
            String key = PARENT_KEYS.peek();
            parentArray.put(parentJson);
            root.getValue().getKey();
            root.getValue().setData(parentArray);
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
