package com.imos.etj;

import java.io.IOException;
import org.json.JSONObject;

/**
 *
 * @author p
 */
public class ExcelToJSONTesting {

    private static String excelFileName;
    private static String excelSheetName;
    private static String jsonFileName;

    public static void main(String[] args) throws IOException {
        excelFileName = "src/main/resources/SampleData.xlsx";
        excelSheetName = "Testcase1";
        excelSheetName = "Testcase2";
        excelSheetName = "Testcase3";
        excelSheetName = "Testcase4";
        excelSheetName = "Testcase5";
        excelSheetName = "Testcase6";
        excelSheetName = "Testcase7";

        jsonFileName = "src/main/resources/testResult7.json";

        //new ExcelToJSON(excelFileName, excelSheetName, jsonFileName).checkInputs(args).generateJSONData().writeToFile();
        String excelFilePath = "/home/p/Documents/Sample.xlsx";
        String sheetName;
        sheetName = "Sheet1";
        sheetName = "Sheet2";
        sheetName = "Sheet3";
        sheetName = "Sheet4";
        sheetName = "Sheet5";
        sheetName = "Sheet6";
        sheetName = "Sheet7";
//        sheetName = "Sheet8";
//        sheetName = "Sheet9";
//        sheetName = "Sheet10";
        JSONObject json = new ExcelDataExtractor().generateJSONDataFromExcelFile(excelFilePath, sheetName);
        System.out.println(json.toString(4));
    }
}
