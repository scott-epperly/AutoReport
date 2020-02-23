# AutoReport
The AutoReport module will allow you to automate the creation of various common file types from data sources like Azure Data Explorer (using KQL) and any SqlClient, OLEDB, ODBC, or OracleClient data source.  The primary purpose is to allow the creation of preformatted documents ready for direct user consumption by supplying only the data source queries, output templates, and a lightweight configuration file.

## Usage
### Export-AutoReport
| Parameter            | Type        |  Req? | Example             |
| ----------------     | ----------- | :---: | ------------------- |
| AutoRptFile          | string      |   Y   | .\analysis1.autorpt |
| TargetConnectionName | string[]    |   Y   | Connection1         |
| OutputPath           | string      |   Y   | c:\temp\analysis1\  |
| UserParamHash        | hashtable[] |   N   | @(@{"myVar1"="val1"},@{"myVar1"="val1";"myVar2"="val2"}) |

Example Usage:
```PowerShell
Export-AutoReport
    -AutoRptFile .\samples\ExcelSample1\ExcelSample1.autorpt
    -TargetConnectionName Connection1
    -OutputPath C:\temp\testout\
    -UserParamHash @(@{"@dbName"="tempdb"}, @{"@dbName"="SSISDB"})
    -Verbose
```
### New-AutoReportPack
| Parameter        | Type        |  Req? | Example             |
| ---------------- | ----------- | :---: | ------------------- |
| SourceFolder     | string      |   Y   | c:\Dev\analysis1\   |
| AutoRptFile      | string      |   Y   | .\analysis1.autorpt |

Example Usage:
```PowerShell
New-AutoReportPack
    -SourceFolder .\samples\ExcelSample1 
    -AutoRptFile .\samples\ExcelSample1.autorpt
    -TargetConnectionName Connection1
    -OutputPath C:\temp\testout\
    -UserParamHash @(@{"@dbName"="tempdb"}, @{"@dbName"="SSISDB"})
    -Verbose
```

## AutoReportPack (.autorpt) Development
You, as the developer, will be creating the .autorpt file to pass to Export-AutoReport.  This file is an archive (zip) file that contains 3 components:
1. config.json
2. templates folder containing your template files
3. Queries folder containing your query files

### config.json
Before diving in, let's review an example config.json.  This lays out the basic structure as well as provides an example of the most common output type, Excel:
```
{
    "Mappings": [
        {
            "OutputFileNameTemplate":"$ArTargetConnection`_$Query`_$($Get-Date()).xlsx",
            "OutputTemplateName":"Template.xlsx"
            "OutputType":"Excel",
            "SourceMaps": [
                {
                    "DataProviderName":"Kusto",
                    "QueryFileName":"MyFirstQuery.kql",
                    "DataTableId":"0",
                    "WorksheetName":"Sheet1",
                    "Cell":"B2",
                    "TableStyle":"Medium2"        
                },
                {
                    "DataProviderName":"Kusto",
                    "QueryFileName":"MySecondQuery.kql",
                    "DataTableId":"0",
                    "WorksheetName":"Sheet2",
                    "Cell":"B2",
                    "TableStyle":"Medium2"
                }
            ]
        }
    ],
    "Connections": [
        {
            "Name":"Connection1",
            "ConnectionString":"Data Source=https://help.kusto.windows.net:443;Initial Catalog=MyCatalog;Federated Security=True"
        },
        {
            "Name":"Connection2",
            "ConnectionString":"Data Source=https://help.kusto.windows.net:443;Initial Catalog=MyCatalog2;Federated Security=True"
        }
    ]
}
```

The timing of the output creation will determine just how often the cmdlet will create an output file.  Generally, each execution of Export-AutoReport will process in the following manner:
```
foreach TargetConnection in TargetConnections
    Connect to the TargetConnection
->  1. Create OutputFile per connection?
    
    foreach ParamHashTable in UserParamHash
->      2. Create OutputFile per UserParamHash?
        
        foreach Query in Queries(apply ParamHashTable)
->          3. Create OutputFile per query?
            QueryResults = Execute Query
            
            foreach QueryResult in QueryResults
->              4. Create OutputFile per result set?
                
                foreach Row in QueryResult
->                  5. Create OutputFile per row? (not yet implemented)
                    Apply Row to OutputFile
                loop
            loop
        loop
    loop
loop
```

You have the ability to decide when to create the file in the process by configuring the OutputFileNameTemplate.  Logically, the deeper in the process you decide to create the file, the greater the probability that more files will be created:

- **Per connection**: To achieve creating a file for every change in iteration, ensure the only dynamic portion of your OutputFileNameTemplate contains `$ARTargetConnection`.
- **Per UserParamHash**: The UserParamHash parameter of Export-AutoReport expects an array of hashtables.  To have a new file be created at this level, you can specify one of your variable names as if it were a PowerShell variable in your OutputFileNameTemplate definition.  For example, if your `-UserParamHash` variable looked like:
```
Export-AutoReport -AutoRptFile test.autorpt -TargetConnection Conn1 -UserParamHash @(@{"myVar"="myVal"})
```
In this case, you could reference $myVar in your OutputFileNameTemplate definition like this:
```
...
"OutputFileNameTemplate":"$myVar.xlsx"
...
```
- **Per Query**: Just like the `$ArTargetConnection`, another variable, `$ArQuery` is made available representing the current query name (note: this will be the name of the query file _without_ the extension).
- **Per Result Set**: An output file will be created for each Connection/ParamHashTable/Query/Result Set combination.  Assuming the query returns data, there is at least 1 Result Set per query.
- **Per Row**: This is yet to be implemented, but should be in the works.

#### Mappings
The Export-AutoReport cmdlet can produce the following types of output.  Each output type has unique properties in the Mappings section that is required for the output to succeed.  Please review the following OutputTypes for specific information.

##### OutputFileNameTemplate
The OutputFileNameTemplate is where you define the Output File name.  Because the process is creating the file dynamically, the file name will need to contain dynamic elements to ensure uniqueness.  In fact, though the previous section outlines the different logical places where a file can be created, the reality is that the OutputFileNameTemplate is actually controlling the behavior.  In general, the following capabilites are available to you:

1. Any PowerShell command that produces a static string
2. User-defined parameters contained in the UserParamHash
3. Variables for `$ARTargetConnection`, `$ARQuery`, `$ARDataTableNumber`, & `$ArRowNumber`

##### SourceMaps
The collection of SourceMaps define which source queries will be executed for the given Mapping output.  Each SourceMap configuration is defined by the OutputType.  However, each SourceMap will contain the following properties:

| Property         | Description
| --------------   | ------------------------------------- |
| DataProviderName | Name of the data provider (see below) |
| QueryFileName    | File name of the query to be executed |
| DataTableId      | The ordinal position of the data table (result set); this is typically 0 |

###### DataProviderName
AutoReport current supports the following DataProviderName values:

| DataProviderName | Data Provider            | Parameter Naming syntax |
| ---------------- | -------------            | ----------------------- |
| Kusto            | n/a                      | Same as defined in KQL - no special characters required |
| SqlClient        | System.Data.SqlClient    | Named parameters: @_name_ |
| Oledb            | System.Data.OleDb        | Positional
| Odbc             | System.Data.Odbc         | Positional
| OracleClient     | System.Data.OracleClient | Named parameters: @_name_


###### OutputType
The OutputType, as the name suggests, defines the type of output file(s) you intend to create.  Below are possible values and their current status of supportability:
| OutputType  | Supported |
| ----------  | :-------: |
| Excel       | Y         |
| Word        | Future    |
| PowerPoint  | Future    |
| FixedWidth  | Future    |
| Delimited   | Future    |
| PDF         | Future    |

###### Excel OutputType
All output to Excel is written into an Excel Table in the mapped Worksheet with the left-most header column starting in the mapped Cell.

| Property       | Description
| -------------- | ------------------------------------- |
| QueryFileName  | File name of the query to be executed |
| DataTableId    | The ordinal position of the data table (result set); this is typically 0 |
| WorksheetName  | Name of the Worksheet (tab) in the Excel template |
| Cell           | The cell location on the worksheet where you want the data to start writing |
| TableStyle     | Name of the style of the table to output |

###### Word OutputType
. . .  coming soon!

###### PowerPoint OutputType
. . .  coming soon!

###### FixedWidth OutputType
. . .  coming soon!

###### Delimited OutputType
. . .  coming soon!

###### PDF OutputType
. . .  coming soon!

##### Connections
The Connections collection represents every possible connection that your AutoReport pack can access.  Rather than passing connection strings when calling Invoke-AutoReport, you will instead reference the `-TargetConnectionName` providing in the AutoReport pack.

For more information regarding Azure Data Explorer (Kusto) connection strings, please review:
https://docs.microsoft.com/en-us/azure/kusto/api/connection-strings/kusto

### Template File
Each OutputType will require a template file in order to apply the data source queries.  Each template type has unique requirements; please review the appropriate Output Type for template guidance:

#### Excel OutputType Templates
Excel is perhaps one of the easiest to create.  Simply format the area(s) around where you intend to land the data table, give tabs appropriate names, and save the file as Template.xls[x|m|b].  Alternatively, you can have a practically blank template workbook and the Worksheet names referenced in the config.json will be created if the Worksheet names don't exist.

### Queries
The last component of the AutoReportPack are the queries that are to be executed to populate the Output File.  As already referenced in the config.json section above, the queries are referenced in the Mappings section to guide the process of populating the Output File.

It's important to note that multiple result sets are supported for both query types; you may choose to implement this approach in your query files.  The DataSetId However, it is generally recommended to limit each query file to just one result set to simplify query code maintenance.

Lastly, most of the Output Types include a header with the data produced.  The header name used will match the name of the column.  Be sure you name the columns with a user-friendly name that accurately describes the data and doesn't heavily rely on unkonwn acronyms and/or source system column names.

#### Working with KQL (.kql or .csl) Files
Query parameters should be specified in the UserParamHash parameter of Invoke-AutoReport.  The parameter names need to match the names in your KQL query files using the `declare query_parameters(myAppName:string, myRequestId:string);`.  For example, let's say you have the following query:

```KQL
declare query_parameters(pLastName:string, pFirstName:string);
MyKustoPersonTable
| where LastName == pLastName
| where FirstName == pFirstName; 
```
Your UserParamHash could look like the following:
```PowerShell
Export-AutoReport
    ...
    -UserParamHash @(@{"pLastName"="Mouse"; "pFirstName"="Mickey"}, @{"pLastName"="Duck"; "pFirstName"="Donald"})
```


#### Working with SQL (.sql) Files
AutoReport supports multiple data providers in order to facilitate a larger range of sources; please see the possible values of DataProviderName in the SourceMaps section.  Note: all but Kusto are considered to be SQL files.
Please reference [Configuring parameters and parameter data types](https://docs.microsoft.com/en-us/dotnet/framework/data/adonet/configuring-parameters-and-parameter-data-types?view=netframework-4.8)

Example: SQL Server using SqlClient
```SQL
SELECT *
FROM dbo.Person
WHERE LastName = @LastName
    and FirstName = @FirstName;
```
Your UserParamHash could look like the following:
```PowerShell
Export-AutoReport
    ...
    -UserParamHash @(@{"@LastName"="Mouse"; "@FirstName"="Mickey"}, @{"@LastName"="Duck"; "@FirstName"="Donald"})
```

## Contributing
Your contributions are welcome!  As you can see, we currently only support output to Excel, but the other types need some love!  Moreover, there's much work to be done to get all of the anticipated source data providers working as expected.

## Version History
### 0.0.1
Initial release.  This release only supports Kusto (including parameters) and SqlClient sources (without parameters) as well as the Excel output type.  There are likely lots of bugs, but it's functional enough for initial purposes (e.g. if you're holding your tongue right and the stars are all aligned)!