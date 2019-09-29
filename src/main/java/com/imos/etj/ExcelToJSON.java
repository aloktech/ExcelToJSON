package com.imos.etj;

import java.io.File;
import java.io.IOException;
import java.lang.reflect.Field;
import java.nio.file.Files;
import java.nio.file.OpenOption;
import java.nio.file.Paths;
import java.nio.file.StandardOpenOption;
import java.util.LinkedHashMap;
import org.json.JSONException;
import org.json.JSONObject;

/**
 *
 * @author p
 */
public class ExcelToJSON {

    public static final OpenOption[] OPEN_OPTION = new OpenOption[]{
        StandardOpenOption.CREATE,
        StandardOpenOption.TRUNCATE_EXISTING};

    private static String excelFileName;
    private static String excelSheetName;
    private static String jsonFileName;
    private final static boolean PROD_MODE = false;

    public static void main(String[] args) throws IOException {
        if (checkInputs(args)) {
            return;
        }
        new ExcelToJSON().buildJSON();
    }

    private static boolean checkInputs(String[] args) {
        if (PROD_MODE) {
            if (args.length != 2) {
                System.out.println("Enter : <Excel File Name>:<Excel Sheet Name> <JSON File name>");
                return true;
            } else {
                if (!args[0].contains(":")) {
                    System.out.println("Enter : <Excel File Name>:<Excel Sheet Name> <JSON File name>");
                    return true;
                }
                String[] data = args[0].split(":");
                String excelFileNameTemp = data[0].trim();
                if (!excelFileNameTemp.endsWith(".xlsx")) {
                    System.out.println("Enter : <Excel File Name>:<Excel Sheet Name> <JSON File name>");
                    return true;
                }
                String excelSheetNameTemp = data[1].trim();
                if (excelSheetNameTemp.isEmpty()) {
                    System.out.println("Enter : <Excel File Name>:<Excel Sheet Name> <JSON File name>");
                    return true;
                }
                excelFileName = excelFileNameTemp;
                excelSheetName = excelSheetNameTemp;
                String jsonFileNameTemp = args[1].trim();
                if (!jsonFileNameTemp.endsWith(".json")) {
                    System.out.println("Enter : <Excel File Name>:<Excel Sheet Name> <JSON File name>");
                    return true;
                }
                jsonFileName = jsonFileNameTemp;
            }
        }
        return false;
    }

    private void buildJSON() {
        if (PROD_MODE) {
            excelFileName = checkFileExist(excelFileName);
        } else {
            excelFileName = "src/main/resources/SampleData.xlsx";
            excelSheetName = "Testcase1";
            excelSheetName = "Testcase2";
            excelSheetName = "Testcase3";
            excelSheetName = "Testcase4";
            excelSheetName = "Testcase5";
            excelSheetName = "Testcase6";
//            excelSheetName = "Testcase7";
            jsonFileName = "testResult.json";
        }
        
        ExcelDataExtractor excelDataExtractor = new ExcelDataExtractor();
        excelDataExtractor.collectDataFromExcelFile(excelFileName, excelSheetName);
        
        JSONGenerator generator = new JSONGenerator();
        JSONTreeNode root = generator.setJSONValue(excelDataExtractor.getJSONKeyValueMap());
        if (root.getValue() == null) {
            JSONObject result = new JSONObject();
            setMapAsLinkedListMap(result);
            generator.buildJSONObject(root, result, null);
            if (PROD_MODE) {
                jsonFileName = checkFileExist(jsonFileName);
            } else {
                jsonFileName = "src/main/resources/testResult.json";
                jsonFileName = checkFileExist(jsonFileName);
            }
            System.out.println(result.toString(4));
            try {
                Files.write(Paths.get(jsonFileName), result.toString(4).getBytes(),
                        StandardOpenOption.TRUNCATE_EXISTING, StandardOpenOption.CREATE);
            } catch (IOException | JSONException e) {
                e.printStackTrace();
            }
        }
    }

    private String checkFileExist(String fileName) {
        File file = new File(fileName);
        if (!file.exists()) {
            fileName = System.getProperty("user.dir") + File.separator + fileName;
        }
        return fileName;
    }

      private void setMapAsLinkedListMap(JSONObject json) {
        try {
            Field map = json.getClass().getDeclaredField("map");
            map.setAccessible(true);//because the field is private final...
            map.set(json, new LinkedHashMap<>());
            map.setAccessible(false);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
