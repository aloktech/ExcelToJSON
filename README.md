# ExcelToJSON
Check the link : https://excel-to-json-generator.herokuapp.com/ to generate JSON from Excel online<br/><br/>

To generate a JSON file from Excel file in local system.
- Install JDK 8/higher or check JDK 8/higher is installed
- Download the jar file **ExcelToJSON.jar** to a folder
- In the console enter: **java -jar ExcelToJSON.jar**, and press enter, then you will see <br/>
Enter : &lt;Excel File Name&gt;:&lt;Excel Sheet Name&gt; &lt;JSON File name&gt; <br/>
**i.e java -jar ExcelToJSON.jar "SampleData.xlsx:Testcase4" testResult1.json** <br/>
where **SampleData.xlsx** is the Excel file and **Testcase1** is the Sheet name <br/><br/>

**Input Sample Excel Data** <br/>
![Sample Excel Data](docs/new_image.png)<br/>
**Output Sample JSON Data**
```
{
    "key0": "value",
    "key1": "",
    "key2": {"value21": "value"},
    "key3": ["value31"],
    "key4": {
        "key41": {"value411": "value"},
        "key42": "value"
    },
    "key5": [{"value51": "value"}],
    "key6": [
        1,
        2
    ],
    "key7": [
        true,
        false
    ],
    "key8": [
        {
            "key81": "value",
            "key82": true,
            "key83": 23.56,
            "key84": "value"
        },
        {
            "key81": "value",
            "key82": true,
            "key83": 23.56,
            "key84": "value"
        }
    ],
    "key9": "value",
    "key10": [1.2],
    "key11": "value"
}
```
