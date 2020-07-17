package com.imos.etj;

import java.io.Console;
import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.OpenOption;
import java.nio.file.Paths;
import java.nio.file.StandardOpenOption;
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
    public static boolean PROD_MODE = false;

    private String excelFileName;
    private String excelSheetName;
    private String jsonFileName;
    private JSONObject result;

    public static void main(String[] args) throws IOException {
        Console console = System.console();
        if (console != null) {
            PROD_MODE = true;
        }
        new ExcelToJSON()
                .checkInputs(args)
                .generateJSONData()
                .writeToFile();
    }

    public ExcelToJSON() {
        this(null, null, null);
    }

    public ExcelToJSON(String excelFileName, String excelSheetName) {
        this(excelFileName, excelSheetName, null);
    }

    public ExcelToJSON(String excelFileName, String excelSheetName, String jsonFileName) {
        this.excelFileName = excelFileName;
        this.excelSheetName = excelSheetName;
        this.jsonFileName = jsonFileName;
    }

    public ExcelToJSON checkInputs(String[] args) {
        if (PROD_MODE) {
            if (args.length != 2) {
                System.out.println("Enter : <Excel File Name>:<Excel Sheet Name> <JSON File name>");
            } else {
                if (!args[0].contains(":")) {
                    System.out.println("Enter : <Excel File Name>:<Excel Sheet Name> <JSON File name>");
                }
                String[] data = args[0].split(":");
                String excelFileNameTemp = data[0].trim();
                if (!excelFileNameTemp.endsWith(".xlsx")) {
                    System.out.println("Enter : <Excel File Name>:<Excel Sheet Name> <JSON File name>");
                }
                String excelSheetNameTemp = data[1].trim();
                if (excelSheetNameTemp.isEmpty()) {
                    System.out.println("Enter : <Excel File Name>:<Excel Sheet Name> <JSON File name>");
                }
                excelFileName = excelFileNameTemp;
                excelSheetName = excelSheetNameTemp;
                String jsonFileNameTemp = args[1].trim();
                if (!jsonFileNameTemp.endsWith(".json")) {
                    System.out.println("Enter : <Excel File Name>:<Excel Sheet Name> <JSON File name>");
                }
                jsonFileName = jsonFileNameTemp;
            }
        }
        return this;
    }

    protected ExcelToJSON generateJSONData() {
        if (PROD_MODE) {
            excelFileName = checkFileExist(excelFileName);
        }

        ExcelDataExtractor excelDataExtractor = new ExcelDataExtractor();
        result = excelDataExtractor.generateJSONDataFromExcelFile(excelFileName, excelSheetName);
        System.out.println(result.toString(4));
        return this;
    }

    protected void writeToFile() {
        if (jsonFileName == null) {
            return;
        }
        try {
            Files.write(Paths.get(jsonFileName), result.toString(4).getBytes(),
                    StandardOpenOption.TRUNCATE_EXISTING, StandardOpenOption.CREATE);
        } catch (IOException | JSONException e) {
            e.printStackTrace();
        }
    }

    private String checkFileExist(String fileName) {
        File file = new File(fileName);
        if (!file.exists()) {
            fileName = System.getProperty("user.dir") + File.separator + fileName;
        }
        return fileName;
    }
}
