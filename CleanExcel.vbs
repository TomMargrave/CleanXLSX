'  Created by :tom.margrave at Orasi Support
'  File updated by Tom Margrave at Qualitest  
'  File created:Fri Mar 24 2017 16:23:02 GMT-0400 (Eastern Daylight Time)
'  File Name  CleanExcel.vbs
''
'  This file will take one input Excel file and do the following:
'   1. Save each sheet as a CSV file.
'   2. Store the order of the sheet(s) in sheet.txt file.
'   3. Open each CSV file and do the following.
'       a. Remove all non-breaking spaces ASCII 160
'       b. Remove extra blank columns on the right side of all rows.
'       c. Remove rows that have no data.
'   4. Save changes to CSV files.
'   5. Convert CSV files to EXCEL files.
'   6. Combine EXCEL files into new EXCEL file with the sheets in same  order as orignal.
'   7. Remove any temporary files created during process.

Dim cDir         ' current working directory base on source file'
Dim cnt          ' counter variable '
Dim DoNotSkip    ' boolean to help recreate tabs in new xlsx'
Dim myArr        ' temporary array to count items'
Dim objExcel     ' Excel object'
Dim objFile      ' used for varous File read and write operations.'
Dim objFilecsv   ' current CSV file object '
Dim objFileSheet ' File object for the sheet'
Dim objFSO		 ' File System Object
Dim objFSOcsv    ' Read and update CSV files after changes. '
Dim objFSOSheet  ' used to open sheet file
Dim objWB        ' Excel workbook object'
Dim objWS        ' Excel work sheet object'
Dim oFSODel      ' file System Object to clean up temporary files.'
Dim outFile      ' holder for temporary file holds sheet names and order.'
Dim sHeaderLog   ' holds the log of any changes to header'
Dim srcOne       ' holds the active source xlsx file when rebuilding doc'
Dim srcTwo       ' holds the active source xlsx file when rebuilding doc'
Dim strCSVFile   ' holds current CSV file name and path'
Dim strFileName  ' holds file name
Dim strLine      ' current sheet processed
Dim strReplace   ' string to replace with'
Dim strSearch    ' string to search for'
Dim strText      ' holds current contents of the current CSV file'
Dim strTotals    ' holds results for current sheet.'
Dim tgtXLSX      ' holds the output file path and name'
Dim totalColumn  ' toltal columns removed from sheet'
Dim totalNBSpace ' total non-breaking space'
Dim totalRow     ' total rows removed from sheet'
Dim bExcelVisible ' boolean to set Excel Visible'
Dim bExcelAlerts   ' boolean to set Excel display Alerts'
Dim sColumnDupe     ' results of duplicate column names

Const ForReading = 1
Const ForWriting = 2

bExcelVisible = False
bExcelAlerts = False
outFile = "Sheet.txt"

Set objFSOSheet = CreateObject("Scripting.FileSystemObject")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set myLog = objFSO.OpenTextFile("CleanExcel.log", ForWriting, True)
myLog.WriteLine(now)
myLog.WriteLine "Started run"

myProcess = 0
cnt = WScript.Arguments.Count

'If third arg set then suppress dailogs
If cnt > 2 then
 supressNotes = 1
End If

'Check to see If Excel is running'
If IsProcessRunning("Excel.exe") Then
    myEcho("Excel is running. " & vbCrLf & " Please close Excel to process." )
    cnt = 0
End If

'Check the source file'
If cnt > 0 Then
    strFileName =  WScript.Arguments.Item(0)
    myLog.WriteLine "Source file:" & strFileName
    If objFSO.FileExists(strFileName) Then
        Set objFile = objFSO.GetFile(strFileName)
        cDir = objFSO.GetParentFolderName(objFile)
        Set objFile = Nothing
        myProcess = 1
    else
        myEcho("Source file does not exist at " & strFileName)
        myProcess = 0
    End If
End If

'Check the secondary attributes'
If cnt > 1 Then
    tgtXLSX = WScript.Arguments.Item(1)
    ' Check to see If extension is xls or xlsx and length greater than 6'
    If(Instr(LCase(Right(tgtXLSX, 5)),".xls") > 0) AND (Len(tgtXLSX) > 6) Then
        myProcess = 1
    Else
        strText = "Second Attribute does not have .xlsx Extension or " & vbCrLf
        strText = strText & "  length is too small"
        myEcho(strText)
        myProcess = 0
    End If
    If(Instr(tgtXLSX,"\") > 1) OR (Instr(tgtXLSX,":") > 1) Then
        myEcho("Second attribute has path information ")
        myEcho(" ")
        myProcess = 0
    End If
Else
    tgtXLSX = "New.xlsx"
End If

myLog.WriteLine "Target file: " & tgtXLSX
'log setting of Supress
If cnt > 2 then
    myLog.WriteLine "Supress Dialogs set"
End If
myLog.WriteLine  vbCrLf


tgtXLSX = cDir & "\" & tgtXLSX

If myProcess = 1 then
    Call CleanUp()
    Call WriteFile(strFileName)
    Call CheckCSV()
    Call CheckColumnHeaders()
    Call CSVtoExcel()
    Call ExcelCombine()
    Call CleanUp()
    If Len(sHeaderLog) > 4 then
        myEcho "# # # # # #" & sHeaderLog & vbCrLf &  "# # # # # #" & vbCrLf
    End If
    If Len(sColumnDupe) > 4 Then
        myEcho "# # # # # #" & sColumnDupe & vbCrLf &  "# # # # # #" & vbCrLf
    End If
    myEcho(strTotals)
Else
    If Not(supressNotes = 1) Then displayHELP()
End If

myLog.Close
Set myLog = Nothing
Set objFSO = Nothing

'#######################################################################
'  ## F U N C T I O N  and S U B below this line ###
'#######################################################################

'**********************************************************************
' Sub Name: ExcelCombine
' Purpose: Combine Excel files into one file name New.xlsx
' Author: Tom Margrave
' Input:
'	None
' Return: None
' Prerequisites:
' var outFile
'     cDir
'**********************************************************************
Sub ExcelCombine()
    'open sheets to parse thru
    Set objFileSheet = objFSOSheet.OpenTextFile(outFile)
    cnt = 0
    Do Until objFileSheet.AtEndOfStream
        DoNotSkip = True
        strLine = objFileSheet.ReadLine
        If(cnt = 0 ) Then
            srcOne = cDir & "\" & strLine
            DoNotSkip = false
            myCopyFile srcOne & ".xlsx",tgtXLSX
        elseIf (cnt > 1) then
            srcOne =  tgtXLSX
        End If

        srcTwo = cDir & "\" &strLine
        If(DoNotSkip = True) then
            On Error Resume Next ' Turn on the error handling flag
            Set objExcel = GetObject(,"Excel.Application")
            'If not found, create a new instance.
            If Err.Number = 429 Then  '> 0
              Set objExcel = CreateObject("Excel.Application")
            End If
            On Error GOTO 0
            ' Set objExcel = CreateObject("Excel.Application")

            objExcel.Visible = bExcelVisible
            objExcel.DisplayAlerts = bExcelAlerts

            Set wbSource = objExcel.Workbooks.Open(srcOne)
            Set objWorkbook1 = objExcel.Workbooks.Open(srcTwo)

            wbSource.Activate

            'wbSource.UsedRange.Select

            objWorkbook1.Sheets.Copy , wbSource.Sheets(1)

            'Get name of sheet without extension '
            arr = Split(wbSource.Sheets(1).Name, ".")
            'rename worksheets
            'wbSource.Sheets(1).Name = arr(0)

            'Get name of sheet without extension
            arr = Split(objWorkbook1.Name, ".")

            'rename worksheets
            wbSource.Sheets(2).Name = arr(0)

            ' Move new sheet to last position
            wbSource.Sheets(2).Move , wbSource.Sheets(wbSource.Sheets.Count)
            'move'
            'wbSource.Sheets(1).Select


            'Save Spreadsheet, 51 = Excel 2007-2010
            wbSource.SaveAs tgtXLSX, 51

            'Release Lock on Spreadsheet
            objExcel.Quit()
            Set ObjExcel = Nothing
            Set wbSource = Nothing
            Set objWorkbook1 = Nothing

        End If
        cnt = cnt+1
        waitExcelStop()
    Loop
    objFileSheet.Close
    Set objFileSheet = Nothing
End Sub

'**********************************************************************
' Sub Name: WriteFile
' Purpose: Converts Excel file to CSV file per sheet
' Author: Tom Margrave
' Input: strFileName  path to the source Excel
'	None
' Return: Nothing but creates CSV file for each sheet
'       outFile  location to store sheet names and order'
' Prerequisites: Excel object
''**********************************************************************
Sub WriteFile(ByVal strFileName)
    Set objFileSheet = objFSOSheet.CreateTextFile(outFile,True)
    Set objExcel = CreateObject("Excel.Application")
    Set objWB = objExcel.Workbooks.Open(strFileName)

    ' cycle thru the sheets''objExcel.Workbooks.Count)
    For Each objWS In objWB.Sheets
        'write sheet name to file for later usage'
        objFileSheet.Write objWS.Name & vbCrLf
        'copy sheet'
        objWS.Copy
        objExcel.ActiveWorkbook.SaveAs objWB.Path & "\" & objWS.Name & ".csv", 6
        objExcel.ActiveWorkbook.Close False
    Next

    objWB.Close False
    objExcel.Quit
    objFileSheet.Close
    Set objWB = Nothing
    Set objExcel = Nothing
    set objFileSheet = Nothing
    waitExcelStop()
End Sub

'**********************************************************************
' Sub Name: CheckColumnHeaders
' Purpose:  Check CSV header for bad characters and duplicates
' Author: Tom Margrave
' Input:
'	None
' Return:
' Prerequisites:
''**********************************************************************
Sub CheckColumnHeaders()
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFileSheet = objFSOSheet.OpenTextFile(outFile)
    Set objFSOcsv = CreateObject("Scripting.FileSystemObject")
    Do Until objFileSheet.AtEndOfStream
        strLine = objFileSheet.ReadLine
        strCSVFile = cDir & "\" & strLine & ".csv "
        iLine = 0
        Set objFilecsv = objFSOcsv.OpenTextFile(strCSVFile, ForReading)
        sFullFile = "~~~"
        Do Until objFilecsv.AtEndOfStream
            strText = objFilecsv.ReadLine
            iLine = iLine + 1
            If iLine = 1 Then
                arrayLine = Split(strText,",")
                strText = "~~~"
                totalColumns = ubound(arrayLine)

                ' For Each sColumn in arrayLine
                For x = 0 to totalColumns
                    sColumn = arrayLine(x)

                    sColumn = checkfor(sColumn, "<")
                    sColumn = checkfor(sColumn, ">")
                    sColumn = checkfor(sColumn, "`")
                    sColumn = checkfor(sColumn, "~")
                    sColumn = checkfor(sColumn, "%")

                    'Look for duplicates column names'
                    sLColumn = LCase(sColumn)
                    For y = x + 1 to totalColumns
                        If (sLColumn = lcase(arrayLine(y))) Then
                            arrayLine(y) = arrayLine(y) & "_DUPE"
                            sColumnDupe = sColumnDupe & vbCrLf & "Sheet: " & strLine
                            sColumnDupe = sColumnDupe & vbCrLf & "Duplicate column: "
                            sColumnDupe = sColumnDupe & vbCrLf & sColumn
                            sColumnDupe = sColumnDupe & vbCrLf & "  Renamed to: "  & arrayLine(y)
                            sColumnDupe = sColumnDupe & vbCrLf
                        End If
                    Next

                    If strText = "~~~" Then
                        strText =  sColumn
                    Else
                        strText = strText & "," & sColumn
                    End If
                Next
            End If

            If sFullFile =  "~~~"  Then
                sFullFile =  strText
            Else
                sFullFile =  sFullFile & vbCrLf & strText
            End If
        Loop 'csv lines'

        objFilecsv.Close

        'write lines back to CSV File'
        Set objFile = objFSO.OpenTextFile(strCSVFile, ForWriting)
        objFile.WriteLine sFullFile
        objFile.Close
        Set objFile = Nothing

    Loop 'sheet'
    objFileSheet.Close
    Set objFilecsv = Nothing
    Set objFSOcsv = Nothing
    Set objFileSheet = Nothing
    Set objFSO = Nothing
End Sub
'**********************************************************************
' Sub Name: checkfor
' Purpose:  Check for sVar input and replace then report results to log file
' Author: Tom Margrave
' Input:
'	None
' Return:
' Prerequisites:
''**********************************************************************
Function checkfor(sVar, sSearch)
    'count of search items
    cnt = countExist(sVar, sSearch)
    'set default return variable
    checkfor = sVar

    If cnt > 0 Then
        sReplace = "_"
        checkfor = Replace(sVar, sSearch, sReplace)
        'Log issue'
        sHeaderLog = sHeaderLog & vbCrLf & "Sheet: " & strLine
        sHeaderLog = sHeaderLog & vbCrLf & "Column: " & sVar
        sHeaderLog = sHeaderLog & vbCrLf & " has character: " & sSearch
        sHeaderLog = sHeaderLog & vbCrLf & " Replaced with: " & checkfor
    End If
End Function



'**********************************************************************
' Sub Name: CheckCSV
' Purpose:  Check and remove Non-breaking spaces, blank rows, and blank columns
'       from CSV file
' Author: Tom Margrave
' Input:
'	None
' Return:
' Prerequisites:
''**********************************************************************
Sub CheckCSV()
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFileSheet = objFSOSheet.OpenTextFile(outFile)
    Set objFSOcsv = CreateObject("Scripting.FileSystemObject")
    Do Until objFileSheet.AtEndOfStream
        strLine = objFileSheet.ReadLine
        strCSVFile = cDir & "\" & strLine & ".csv "

        Set objFilecsv = objFSOcsv.OpenTextFile(strCSVFile, ForReading)
        strText = objFilecsv.ReadAll
        objFilecsv.Close
        totalColumn = 0
        totalRow = 0
        totalNBSpace = 0

        'set the search and replace to remove extra column'
        strReplace = "," & Chr(13) & Chr(10)
        strSearch = ",," & Chr(13) & Chr(10)
        cnt = 1

        Do Until cnt = 0
            cnt = countExist(strText, strSearch)
            strText = Replace(strText, strSearch, strReplace)
            totalColumn = totalColumn + cnt
        Loop

        ' Look for empty rows
        strReplace = "" & Chr(13) & Chr(10)
        strSearch = Chr(13) & Chr(10) & "," & Chr(13) & Chr(10)
        cnt = countExist(strText, strSearch)
        strText = Replace(strText, strSearch, strReplace)
        totalRow = cnt

        ' Look for non-breaking spaces
        strReplace = " "
        strSearch = Chr(160)
        cnt = countExist(strText, strSearch)
        strText = Replace(strText, strSearch, strReplace)
        totalNBSpace = cnt

        Set objFile = objFSO.OpenTextFile(strCSVFile, ForWriting)
        objFile.WriteLine strText

        objFile.Close
        Set objFile = Nothing
        strTotals  = strTotals & "Sheet:  " & strLine & vbCrLf
        strTotals  = strTotals & "Total Columns Cleaned:  " & totalColumn  & vbCrLf
        strTotals  = strTotals & "Total Rows Cleaned:  " & totalRow & vbCrLf
        strTotals  = strTotals & "Total Non-breaking Spaces Cleaned:  " & totalNBSpace   & vbCrLf
        strTotals  = strTotals & vbCrLf
    Loop
    objFileSheet.Close
    Set objFileSheet = Nothing
End Sub

'**********************************************************************
' Sub Name: CSVtoExcel
' Purpose:  Converts CSV files into  Excel Document per csv file
' Author: Tom Margrave
' Input:
'	None
' Return:  Nothing  but output xlsx files for each CSV file
' Prerequisites:
'**********************************************************************
Sub CSVtoExcel()
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFileSheet = objFSO.OpenTextFile(outFile)

    Do Until objFileSheet.AtEndOfStream
        strLine = objFileSheet.ReadLine

        srcCSVFile =  cDir & "\" & strLine & ".csv "
        tgtXLSFile = cDir & "\" & strLine & ".xlsx"

        'Create Spreadsheet
        Set objExcel = CreateObject("Excel.Application")

        objExcel.Visible = bExcelVisible
        objExcel.DisplayAlerts = bExcelAlerts

        'Import CSV into Spreadsheet
        Set objWorkbook = objExcel.Workbooks.Open(srcCSVFile)
        Set objWorksheet1 = objWorkbook.Worksheets(1)

        'Adjust width of columns
        'This causes errors on some machines and commented out.'
        ' Set objRange = objWorksheet1.UsedRange
        ' objRange.EntireColumn.Autofit
        objWorksheet1.SaveAs tgtXLSFile, 51

        'Release Lock on Spreadsheet
        objExcel.Quit()
        Set objWorksheet1 = Nothing
        Set objWorkbook = Nothing
        Set ObjExcel = Nothing
        waitExcelStop()
    Loop

    objFileSheet.Close
    set objFileSheet = Nothing
End Sub

'**********************************************************************
' Sub Name: myDelFile
' Purpose: Delete the file
' Author: Tom Margrave
' Input:
'	myFile  path to file to be deleted
' Return:  None
' Prerequisites:
''**********************************************************************
Sub myDelFile(myFile)
    Set oFSODel = CreateObject("Scripting.FileSystemObject")
    If oFSODel.FileExists(myFile) Then
        oFSODel.DeleteFile myFile, True
    Else
        myEcho("The file does not exist." & myFile)
    End If
    Set oFSODel = Nothing
End Sub

'**********************************************************************
' Sub Name: CleanUp
' Purpose:  Remove extra files created duing the running of this script.
' Author: Tom Margrave
' Input:
'	None
' Return:
' Prerequisites:
'           outFile   holds listing of files.
'**********************************************************************
Sub CleanUp()
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    If objFSO.FileExists(outFile) Then
        Set objFileSheet = objFSO.OpenTextFile(outFile)
        Do Until objFileSheet.AtEndOfStream
            strLine = cDir & "\" & objFileSheet.ReadLine
            myDelFile(strLine & ".csv")
            myDelFile( strLine & ".xlsx")
        Loop
        objFileSheet.Close
        myDelFile(cDir & "\Sheet.txt")'
        Set objFileSheet = Nothing
    End If
    Set objFSO = Nothing
End Sub

'**********************************************************************
' Sub Name: countExist
' Purpose:  get the total number of sub strings in string
' Author: Tom Margrave
' Input:
'	strValue   String to be searched
'	strSearch   Sub string to be found in strValue
' Return: integer total number of sub strings found
' Prerequisites:
'**********************************************************************
Function countExist(strValue,strSearch)
    myArr = Split(strValue,strSearch)
    countExist = Ubound(myArr)
End Function

'**********************************************************************
' Sub Name: displayHELP
' Purpose:  Display help for this file
' Author: Tom Margrave
' Input:
'	none
' Return: None
' Prerequisites:
'**********************************************************************

Function displayHELP()
	strTemp = strTemp & vbCRLF & "This is a tool to clean Excel files. This tool will do the following:"
	strTemp = strTemp & vbCRLF & " "
	strTemp = strTemp & vbCRLF & "1. Read source Excel file."
	strTemp = strTemp & vbCRLF & "2. Save each sheet as a .CSV file."
	strTemp = strTemp & vbCRLF & "3. Open each .CSV file and do the following:"
	strTemp = strTemp & vbCRLF & vbTAB & "a. Remove all non-breaking spaces: ASCII 160"
	strTemp = strTemp & vbCRLF & vbTAB & "b. Remove extra blank columns on the right side of all rows."
	strTemp = strTemp & vbCRLF & vbTAB & "c. Remove rows that have no data."
	strTemp = strTemp & vbCRLF & "4. Convert .CSV files to target EXCEL file."
	strTemp = strTemp & vbCRLF
	strTemp = strTemp & vbCRLF & "syntax:"
	strTemp = strTemp & vbCRLF & "    cleanExcel.vbs ""path\<source file>"" [<target file> [<suppress>]] "
	strTemp = strTemp & vbCRLF
	strTemp = strTemp & vbCRLF & vbTAB & "<source file>" & vbTAB & "Required - Path with file name to be"
	strTemp = strTemp & vbCRLF & vbTAB & vbTAB & vbTAB &  "    processed. Enclosed in double quotes("")."
	strTemp = strTemp & vbCRLF
	strTemp = strTemp & vbCRLF & vbTAB & "<taget file>" & vbTAB & "Optional - Name of the new Excel file."
	strTemp = strTemp & vbCRLF
	strTemp = strTemp & vbCRLF & vbTAB & "<suppress>" & vbTAB & "Optional - If value is set to '1', all dialogs"
	strTemp = strTemp & vbCRLF & vbTAB & vbTAB & vbTAB & "    will be suppressed."
	strTemp = strTemp & vbCRLF
	myEcho(strTemp)
End Function

'**********************************************************************
' Sub Name: myEcho
' Purpose:  Display messages depending on flag supressNotes
' Author: Tom Margrave
' Input:
'	supressNotes
' Return: None
' Prerequisites:
'**********************************************************************
Function myEcho(strTemp)
    If Not(supressNotes = 1) Then
        WScript.Echo strTemp
    End If
    myLog.WriteLine strTemp

End Function

'**********************************************************************
' Sub Name: IsProcessRunning
' Purpose:  Check to see if Excel is running before starting
' Author: Tom Margrave
' Input:
'	none
' Return: None
' Prerequisites:
'**********************************************************************
Function IsProcessRunning( strProcess)
    strComputer = "."
    Dim Process, strObject
    IsProcessRunning = False
    strObject = "winmgmts://" & strComputer
    For Each Process in GetObject(strObject).InstancesOf("win32_process")
    If UCase(Process.name) = UCase(strProcess) Then
        IsProcessRunning = True
        Exit Function
    End If
    Next
End Function

'**********************************************************************
' Sub Name: waitExcelStop
' Purpose:  On some systems, Excel does not close fast enough.  This
'   will check to see if Excel is running before going forward
' Author: Tom Margrave
' Input:
'	none
' Return: None
' Prerequisites:  Method IsProcessRunning()
'**********************************************************************
Function waitExcelStop()
'Check to see If Excel is running'

    If IsProcessRunning("Excel.exe") Then
        bRun = True
        Do While bRun = True
                ' body
            If NOT(IsProcessRunning("Excel.exe")) Then
                bRun = False
            else
                WScript.Sleep(1000)
                sleepCnt = sleepCnt+1
            End If
            If sleepCnt > 100 Then
                WScript.Echo "Excel is running and will not close. Exit VBS"
                WScript.Quit
            End If
        Loop
    End If
    'body
End Function

 '**********************************************************************
 '  Function Name: myCopyFile
 '  Purpose: Copy one file to anohter name./location
 '  Author: Tom Margrave
 '  Input:
 '      src source file
 '      tgt target file
 '  Return: None
 '  Prerequisites:
 '**********************************************************************
Function myCopyFile(src, tgt)
    Dim FSO
    ' TODO This should have checking to see if file exist.'
    Set FSO = CreateObject("Scripting.FileSystemObject")
    FSO.CopyFile src, tgt
    Set FSO = Nothing

End Function
