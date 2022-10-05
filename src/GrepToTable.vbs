Option Explicit

Function varMappingOptions(ByVal strFpSettings)
    ' <MappingOptions>
    '   <Encoding type="input">shift-jis</Encoding>
    '   <Encoding type="output">shift-jis</Encoding>
    '   <Search type="recursive">False</Search>
    Dim objDOM, varReturn
    Set objDOM = CreateObject("Msxml2.DOMDocument.6.0")
    objDOM.Load strFpSettings
    ReDim varReturn(5)
    varReturn(0) = objDOM.SelectSingleNode("/Root/MappingOptions/Encoding [@type='input']").Text
    varReturn(1) = objDOM.SelectSingleNode("/Root/MappingOptions/Encoding [@type='output']").Text
    varReturn(2) = objDOM.SelectSingleNode("/Root/MappingOptions/Search [@type='recursive']").Text
    varReturn(3) = objDOM.SelectSingleNode("/Root/MappingOptions/Sort [@type='date']").Text
    varReturn(4) = objDOM.SelectSingleNode("/Root/MappingOptions/Filter [@type='regexp']").Text
    varReturn(5) = Chr(CInt(objDOM.SelectSingleNode("/Root/MappingOptions/ColSep [@type='number']").Text))
    varMappingOptions = varReturn
    Set objDOM = Nothing
End Function

Function varGrepPattern(ByVal strFpSettings, ByVal strPatternIndex)
    ' <GrepPatterns>
    '     <GrepPatterns index="2" name="Greetings">
    '         <ColumnHeader>TimeOfDay FirstName</ColumnHeader>
    '         <Pattern>Good (Morning|Afternonn|Evening), ([a-zA-Z]+)!</Pattern>
    '         <GroupCount>2</GroupCount>
    '         <FileName>GreetingExtract.txt</FileName>
    Dim objDOM, objGP, varReturn
    Set objDOM = CreateObject("Msxml2.DOMDocument.6.0")
    objDOM.Load strFpSettings
    Set objGP = objDOM.SelectSingleNode("/Root/GrepPatterns/GrepPattern [@index='" & strPatternIndex & "']")
    ReDim varReturn(3)
    varReturn(0) = objGP.SelectSingleNode("./ColumnHeader").Text
    varReturn(1) = objGP.SelectSingleNode("./Pattern").Text
    varReturn(2) = CInt(objGP.SelectSingleNode("./GroupCount").Text)
    varReturn(3) = objGP.SelectSingleNode("./FileName").Text
    varGrepPattern = varReturn
    SEt objGP = Nothing
    Set objDOM = Nothing
End Function

''' FPath.vbs
Function varFilePropArray(ByRef objFile)
	varFilePropArray = Array( _
	objFile.Path, _
	objFile.Name, _
	objFile.DateLastModified, _
	objFile.Size)
End Function

Const sconFilePropertyArrayPath = 0
Const sconFilePropertyArrayName = 1
Const sconFilePropertyArrayDateLastModified = 2
Const sconFilePropertyArraySize = 3

Function varFileAttrArrayFilteredByPath(ByRef varFileAttr, ByVal strPattern)
	Dim objRgx: Set objRgx = New RegExp
	With objRgx
		.Pattern = strPattern
		.Global = False	' should not be global for file path pattern check
		.IgnoreCase = False	' windows file path system is case insensitive
	End With
	Dim varReturn: ReDim varReturn(-1)
	Dim lngIdxOrg, lngIdxNew
	lngIdxNew = -1
	For lngIdxOrg = LBound(varFileAttr) To UBound(varFileAttr)
		If objRgx.Test(varFileAttr(lngIdxOrg)(sconFilePropertyArrayPath)) Then
			lngIdxNew = lngIdxNew + 1
			ReDim Preserve varReturn(lngIdxNew)
			varReturn(lngIdxNew) = varFileAttr(lngIdxOrg)
		End if
	Next
	varFileAttrArrayFilteredByPath = varReturn
End Function

Sub SortFileAttrArray(ByRef varFileAttr, ByVal intCol)
	Dim i, j, intSwapCount, varTempAttr
	For i = LBound(varFileAttr) + 1 To UBound(varFileAttr)
		intSwapCount = 0
		For j = LBound(varFileAttr) + 1 To UBound(varFileAttr)
			If varFileAttr(j-1)(intCol) > varFileAttr(j)(intCol) Then
				varTempAttr = varFileAttr(j-1)
				varFileAttr(j-1) = varFileAttr(j)
				varFileAttr(j) = varTempAttr
				intSwapCount = intSwapCount + 1
			End If
		Next
		If intSwapCount = 0 Then Exit For
	Next
End Sub

Sub MapFilesTopOnly(ByVal strRootDir, ByRef varFileAttr)
	ReDim varFileAttr(-1)
	Dim lngUB
	Dim objFSO: Set objFSO = CreateObject("Scripting.FileSystemObject")
	Dim objRootDir: Set objRootDir = objFSO.GetFolder(strRootDir)
	Dim objFile
	For Each objFile In objRootDir.Files
		lngUB = UBound(varFileAttr) + 1
		ReDim Preserve varFileAttr(lngUB)
		varFileAttr(lngUB) = varFilePropArray(objFile)
	Next
	Set objFile = Nothing
	Set objRootDir = Nothing
	Set objFSO = Nothing
End Sub

Sub MapFilesRecursivelySub(ByRef objDir, ByRef varFileAttr)
	Dim lngUB
	' Get file attributes.
	Dim objFile: For Each objFile In objDir.Files
		lngUB = UBound(varFileAttr) + 1
		ReDim Preserve varFileAttr(lngUB)
		varFileAttr(lngUB) =  varFilePropArray(objFile)
	Next
	' Call recursively on subfolders.
	Dim objSubDir: For Each objSubDir In objDir.SubFolders
		Call MapFilesRecursivelySub(objSubDir, varFileAttr)
	Next
	Set objFile = Nothing
	Set objSubDir = Nothing
End Sub

Sub MapFilesRecursively(ByVal strRootDir, ByRef varFileAttr)
	ReDim varFileAttr(-1)
	Dim objFSO: Set objFSO = CreateObject("Scripting.FileSystemObject")
	Dim objRootDir: Set objRootDir = objFSO.GetFolder(strRootDir)
	Call MapFilesRecursivelySub(objRootDir, varFileAttr)
	Set objRootDir = Nothing
	Set objFSO = Nothing
End Sub

''' TestFile.vbs
Function strReadText(ByVal strFp, ByVal strCharSet)
	Dim objStream: Set objStream = CreateObject("ADODB.Stream")
	With objStream
		.CharSet = strCharSet
		.Open
		.LoadFromFile(strFp)
		strReadText = .ReadText()
	End With
	Set objStream = Nothing
End Function

Sub WriteText(ByVal strFp, ByVal strTxt, ByVal strCharSet)
	Dim objStream: Set objStream = CreateObject("ADODB.Stream")
	With objStream
		.CharSet = strCharSet
		.Open
		.WriteText strTxt
		.SaveToFile strFp
	End With
	Set objStream = Nothing
End Sub

''' String.vbs
Function strPadZero8(ByVal intNum)
	strPadZero8 = Right("00000000" & CStr(intNum), 8)
End Function

''' Grep.vbs
Function varMatchesSingleGlobal2(ByVal strData, ByRef objRgx, ByVal intGroupUBound)
    Dim objMC, intMcIdx, intGrpCnt, varReturn, intGrpIdx, varRecord
    Set objMC = objRgx.Execute(strData)    
    ReDim varReturn(objMC.Count - 1)
    If objMC.Count > 0 Then
        For intMcIdx = 0 To objMC.Count - 1
            ReDim varRecord(intGroupUBound)
            For intGrpIdx = 0 To intGroupUBound
                varRecord(intGrpIdx) = objMC(intMcIdx).SubMatches(intGrpIdx)
            Next
            varReturn(intMcIdx) = varRecord ' Append record
        Next
    End If
    varMatchesSingleGlobal2 = varReturn ' Return array
End Function

Function objRgxGlobalGroup(ByVal strPattern)
    Dim objRgx: Set objRgx = New RegExp
    With objRgx
        .Pattern = strPattern
        .Global = True
    End With
    Set objRgxGlobalGroup = objRgx
    Set objRgx = Nothing
End Function

Function strArrayTableToString3(ByRef varTbl, ByVal strSepCol)
	Dim intRec, varLines
	ReDim varLines(UBound(varTbl))
	For intRec = LBound(varTbl) To UBound(varTbl)
		varLines(intRec) = Join(varTbl(intRec), strSepCol)
	Next
	strArrayTableToString3 = Join(varLines, vbCrlf)
End Function

''' BinaryFileMerge.vbs
Sub MergeFiles(ByVal strDirIn, ByVal strDirOut, ByVal strFName)
    const adTypeBinary = 1
    ' Input stream
	Dim objStreamIn: Set objStreamIn = WScript.CreateObject("ADODB.Stream")
	objStreamIn.Open
	objStreamIn.type = adTypeBinary
    ' output stream
	Dim objStreamOut: Set objStreamOut = WScript.CreateObject("ADODB.Stream")
	objStreamOut.Open
	objStreamOut.type = adTypeBinary
    ' Loop
	Dim objFSO: Set objFSO = CreateObject("Scripting.FileSystemObject")
	Dim objDirIn: Set objDirIn = objFSO.GetFolder(strDirIn)
	Dim objFile: For Each objFile In objDirIn.Files
        objStreamIn.LoadFromFile(objFile.Path)
        objStreamOut.Write = objStreamIn.Read()
	Next
    ' Save to File
    objStreamOut.SaveToFile(strDirOut & "\" & strFName)
    ' Garbage collection
    Set objStreamIn = Nothing
    Set objStreamOut = Nothing
    Set objDirIn = Nothing
    Set objFile = Nothing
    Set objFSO = Nothing
End Sub

Sub Main()
    ' Get parameters from run.bat
    Dim strDirIn: strDirIn = WScript.Arguments(0)
    Dim strDirOut: strDirOut = WScript.Arguments(1)
    Dim strPatternIndex: strPatternIndex = WScript.Arguments(2)
    Dim strFpSettings: strFpSettings = WScript.Arguments(3)
    Dim strDirTmp: strDirTmp = WScript.Arguments(4)
    WScript.Echo "strDirIn=" & strDirIn
    WScript.Echo "strDirOut=" & strDirOut
    WScript.Echo "strPatternIndex=" & strPatternIndex
    WScript.Echo "strFpSettings=" & strFpSettings
    WScript.Echo "strDirTmp=" & strDirTmp

    ' Get mapping options
    Dim varMapOpt: varMapOpt = varMappingOptions(strFpSettings)
    Const conEncodingInput = 0
    Const conEncodingOutput = 1
    Const conSearchRecursive = 2
    Const conSortByDate = 3
    Const conFilterRegExp = 4
    Const conColSep = 5
    WScript.Echo "varMapOpt(conEncodingInput)=" & varMapOpt(conEncodingInput)
    WScript.Echo "varMapOpt(conEncodingOutput)=" & varMapOpt(conEncodingOutput)
    WScript.Echo "varMapOpt(conSearchRecursive)=" & varMapOpt(conSearchRecursive)
    WScript.Echo "varMapOpt(conSortByDate)=" & varMapOpt(conSortByDate)
    WScript.Echo "varMapOpt(conFilterRegExp)=" & varMapOpt(conFilterRegExp)
    WScript.Echo "varMapOpt(conColSep)='" & varMapOpt(conColSep) & "'"  ' Add ending apos to visualize tab.

    ' Get grep pattern
    Dim varGrpPtn: varGrpPtn = varGrepPattern(strFpSettings, strPatternIndex)
    Const conColumnHeader = 0
    Const conPattern = 1
    Const conGroupCount = 2
    Const conFileName = 3
    WScript.Echo "varGrpPtn(conColumnHeader)=" & varGrpPtn(conColumnHeader)
    WScript.Echo "varGrpPtn(conPattern)=" & varGrpPtn(conPattern)
    WScript.Echo "varGrpPtn(conGroupCount)=" & varGrpPtn(conGroupCount)
    WScript.Echo "varGrpPtn(conFileName)=" & varGrpPtn(conFileName)

    ' Get file path array
    Dim varFileAttr
    ' Populte array (recursively or top only)
    Select Case varMapOpt(conSearchRecursive)
    Case "True"
        WScript.Echo "Search: Recursively"
        MapFilesRecursively strDirIn, varFileAttr
    Case "False"
        WScript.Echo "Search: Top Only"
        MapFilesTopOnly strDirIn, varFileAttr
    Case Else
        WScript.Echo "varMapOpt(conSearchRecursive)=" & varMapOpt(conSearchRecursive)
        Err.Raise vbObjectError + 5, "Setting Value Error", "Value Not Defined"
    End Select
    ' Filter by regular expression
    If Len(varMapOpt(conFilterRegExp)) > 0 Then
        varFileAttr = varFileAttrArrayFilteredByPath(varFileAttr, varMapOpt(conFilterRegExp))
        WScript.Echo "Filter by RegExp: Executed"
    Else
        WScript.Echo "Filter by RegExp: Skipped"
    End If
    ' Sort by date
    Select Case varMapOpt(conSortByDate)
    Case "True"
        SortFileAttrArray varFileAttr, sconFilePropertyArrayDateLastModified
        WScript.Echo "Sort by date: Executed"
    Case "False"
        WScript.Echo "Sort by date: Skipped"
    Case Else
        WScript.Echo "varMapOpt(conSortByDate)=" & varMapOpt(conSortByDate)
        Err.Raise vbObjectError + 5, "Setting Value Error", "Value Not Defined"
    End Select
    WScript.Echo "Count of varFileAttr=" & UBound(varFileAttr) + 1

    ' Prepare the reuseable regular expression object
    Dim objRgx: Set objRgx = objRgxGlobalGroup(varGrpPtn(conPattern))

    ' Write header
    Dim strHeader 
    ' Split(expression[,delimiter[,count[,compare]]])
    strHeader = Join(Split(varGrpPtn(conColumnHeader)), varMapOpt(conColSep)) & vbCrlf
    ' WriteText(ByVal strFp, ByVal strTxt, ByVal strCharSet)
    WriteText strDirTmp & "\" & strPadZero8(0) & ".txt", strHeader, varMapOpt(conEncodingOutput)

    ' Loop through files
    Dim lngFpIdx, strTextIn, varTbl, strTextOut, intGroupUBound
    intGroupUBound = varGrpPtn(conGroupCount) - 1
    For lngFpIdx = LBound(varFileAttr) To UBound(varFileAttr)
        WScript.Echo "lngFpIdx=" & lngFpIdx
        strTextIn = strReadText(varFileAttr(lngFpIdx)(sconFilePropertyArrayPath), varMapOpt(conEncodingInput))
        varTbl = varMatchesSingleGlobal2(strTextIn, objRgx, intGroupUBound)
        ' Function strArrayTableToString3(ByRef varTbl, ByVal strSepCol)
        strTextOut = strArrayTableToString3(varTbl, varMapOpt(conColSep))
        ' WriteText(ByVal strFp, ByVal strTxt, ByVal strCharSet)
        If Len(strTextOut) > 0 Then
            WriteText strDirTmp & "\" & strPadZero8(lngFpIdx + 1) & ".txt", _
            strTextOut & vbCrlf, _
            varMapOpt(conEncodingOutput)
        End If
    Next

    ' Merge Files
    ' Sub MergeFiles(ByVal strDirIn, ByVal strDirOut, ByVal strFName)
    MergeFiles strDirTmp, strDirOut, varGrpPtn(conFileName)
End Sub
Main