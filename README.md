# GrepToTable

## Feature
- Grep text files with regular expression with groups.
- Grouped contents are stored in a table format.

## Development Note
- Only Batch and VBScript are used so that the script can run on a bare windows environment.
- For performance:
    - string concatenation is not in use.
    - dynamic array is initialized based on RegExp match collection.
    - Matched items are exported for file and all files are merged in binary mode.

## How to use
- Execute run.bat from command promopt.
- Input RegExp pattern (*saved on Settings.xml*)
- Input the folder with data files.
- Input the output folder where the result will be stored.
```
+----------------------------------------------------------------------+
| ...\GrepToTable>run                                                  | Execute run.bat
| --- Clean temp                                                       |
| ---Print Options                                                     |
| PrintOptions.vbs: strFpSetting=.\src\Settings.xml                    |
| 1       ShortDescription                                             |
| 2       Greetings                                                    |
| ---Pattern Index Selection                                           |
| Pattern Index: 2                                                     | Select search pattern by index
|                                                                      |
| ---Get Input Folder                                                  |
| Folder with data files: .\data                                       | Set data folder
| DirOut=.\data                                                        |
|                                                                      |
| ---Get output Folder                                                 |
| Folder where output file is stored: C:\_temp                         | Set output folder
| DirOut=C:\_temp                                                      |
|                                                                      |
| ---Run GrepToTable.vbs                                               |
... omit ...
```
## How to setup
- For file mapping option:
    - Set encoding for file reading and file writing.
    - Set search type (either top only or recursive).
    - Set whetehr to sort files by date.
    - Set the output column separator.
- For grep pattern
    - Set index and name: index and name are referenced for which pattern to be used.
    - Set column header separated by spaces
    - Set regular expressoin pattern
    - Set the number of groups in the search pattern
    - Set filter condition for file path in regular expression.
    - Set output file name    
```xml
+------------------------------------------------------------------------------------------------------------+
| <?xml version="1.0" encoding="utf-8"?>                                                                     |
| <Root>                                                                                                     |
|     <ConsoleTitle>Name of Command Prompt Title</ConsoleTitle>                                              |
|     <MappingOptions>                                                                                       |
|         <!-- encoding: (shift-jis, ascii, utf-8) whatever is supported with ADODB.Stream CharSet -->       |
|         <Encoding type="input">shift-jis</Encoding>                                                        |
|         <Encoding type="output">shift-jis</Encoding>                                                       |
|         <!-- recursive search: (True: Recursive, False: Top Only)-->                                       |
|         <Search type="recursive">True</Search>                                                             |
|         <!-- sort by date: (True: Sort by date, False: Do not sort by date)-->                             |
|         <Sort type="date">True</Sort>                                                                      |
|         <!-- filter by regular expression (no filtering in case of empty)-->                               |
|         <Filter type="regexp">.*txt</Filter>                                                               |
|         <!-- column separator(9=tab, 44=comma, 59=semicolon) Converted to a character by function Chr()--> |
|         <ColSep type="number">9</ColSep>                                                                   |
|     </MappingOptions>                                                                                      |
|     <GrepPatterns>                                                                                         |
|         <GrepPattern index="2" name="Greetings">                                                           |
|             <ColumnHeader>TimeOfDay FirstName</ColumnHeader>                                               |
|             <Pattern>Good (Morning|Afternoon|Evening), ([a-zA-Z]+)!</Pattern>                              |
|             <GroupCount>2</GroupCount>                                                                     |
|             <FileName>GreetingExtract.txt</FileName>                                                       |
|         </GrepPattern>                                                                                     |
|     </GrepPatterns>                                                                                        |
| </Root>                                                                                                    |
+------------------------------------------------------------------------------------------------------------+

```