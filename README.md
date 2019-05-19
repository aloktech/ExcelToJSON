# ExcelToJSON

Its a tool to generate a JSON file from Excel data.


- Clone the project
- Build the project in maven(mvn clean install)
- Copy the file **ExcelToJSON-jar-with-dependencies.jar** from target folder to a folder
- In the console enter: **java -jar ExcelToJSON-jar-with-dependencies.jar**, then you will see <br/>
Enter : &lt;Excel File Name&gt;:&lt;Excel Sheet Name&gt; &lt;JSON File name&gt; <br/>
or<br/>
**java -jar ExcelToJSON-jar-with-dependencies.jar "SampleData.xlsx:Testcase4" testResult1.json** <br/>
where **SampleData.xlsx** is the Excel file and **Testcase1** is the Sheet name <br/><br/>

Sample Excel Data <br/>
[Sample Excel Data](docs/Excel File data.png)
```
{
    "sample1": {
        "sample11": 6786868,
        "sample12": {
            "sample121": "Manish",
            "sample122": "Dua"
        },
        "sample13": true
    },
    "sample3": "Testing",
    "sample2": [
        12312,
        123
    ]
}
```
