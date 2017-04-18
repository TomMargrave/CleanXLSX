'  Created by :tom.margrave at Orasi Support
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

Const ForReading = 1
Const ForWriting = 2
outFile="Sheet.txt"

Set objFSOSheet = CreateObject("Scripting.FileSystemObject")
Set objFSO = CreateObject("Scripting.FileSystemObject")

myProcess = 0
cnt = WScript.Arguments.Count

'If third arg set then suppress dailogs
If cnt > 2 then
 supressNotes = 1
End If

'Check to see If Excel is running'
If IsProcessRunning(".", "Excel.exe") Then
    myEcho("Excel is running. " & vbCrLf & " Please close Excel to process." )
    cnt=0
End If

'Check the source file'
If cnt > 0 Then
    strFileName =  WScript.Arguments.Item(0)
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
        strText ="Second Attribute does not have .xlsx Extension or " & vbCrLf
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
    tgtXLSX= "New.xlsx"
End If

tgtXLSX = cDir & "\" & tgtXLSX

If myProcess = 1 then
    Call CleanUp()
    Call WriteFile(strFileName)
    Call CheckCSV()
    Call CSVtoExcel()
    Call ExcelCombine()
    Call CleanUp()
    myEcho(strTotals)
Else
    If Not(supressNotes=1) Then displayHELP()
End If

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
    cnt=0
    Do Until objFileSheet.AtEndOfStream
        DoNotSkip=True
        strLine= objFileSheet.ReadLine
        If(cnt=0 ) Then
            srcOne= cDir & "\" & strLine
            DoNotSkip=false
        elseIf (cnt > 1) then
            srcOne=  tgtXLSX
        end If

        srcTwo= cDir & "\" &strLine
        If(DoNotSkip=True) then
            On Error Resume Next ' Turn on the error handling flag
            Set objExcel = GetObject(,"Excel.Application")
            'If not found, create a new instance.
            If Err.Number = 429 Then  '> 0
              Set objExcel = CreateObject("Excel.Application")
            End If
            On Error GOTO 0
            ' Set objExcel = CreateObject("Excel.Application")

            objExcel.Visible = false
            objExcel.DisplayAlerts=false

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

        end If
        cnt=cnt+1
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
        strLine= objFileSheet.ReadLine
        strCSVFile = cDir & "\" & strLine & ".csv "

        Set objFilecsv = objFSOcsv.OpenTextFile(strCSVFile, ForReading)
        strText = objFilecsv.ReadAll
        objFilecsv.Close
        totalColumn=0
        totalRow=0
        totalNBSpace=0

        'set the search and replace to remove extra column'
        strReplace ="," & Chr(13) & Chr(10)
        strSearch = ",," & Chr(13) & Chr(10)
        cnt=1

        Do Until cnt=0
            cnt=countExist(strText, strSearch)
            strText = Replace(strText, strSearch, strReplace)
            totalColumn = totalColumn + cnt
        Loop

        strReplace ="" & Chr(13) & Chr(10)
        strSearch = Chr(13) & Chr(10) & "," & Chr(13) & Chr(10)
        cnt=countExist(strText, strSearch)
        strText = Replace(strText, strSearch, strReplace)
        totalRow = cnt

        strReplace =" "
        strSearch = Chr(160)
        cnt=countExist(strText, strSearch)
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
        strLine= objFileSheet.ReadLine

        srcCSVFile =  cDir & "\" & strLine & ".csv "
        tgtXLSFile = cDir & "\" & strLine & ".xlsx"

        'Create Spreadsheet
        Set objExcel = CreateObject("Excel.Application")

        objExcel.Visible = false
        objExcel.DisplayAlerts= false

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
            strLine= cDir & "\" & objFileSheet.ReadLine
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
    strTemp = strTemp & vbCRLF & "This is a tool to clean Excel files. This tool will do the following."
    strTemp = strTemp & vbCRLF & " "
    strTemp = strTemp & vbCRLF & "1. Read source Excel file."
    strTemp = strTemp & vbCRLF & "2. Save each sheet as a CSV file."
    strTemp = strTemp & vbCRLF & "3. Open each CSV file and do the following."
    strTemp = strTemp & vbCRLF & "       a. Remove all non-breaking spaces ASCII 160"
    strTemp = strTemp & vbCRLF & "       b. Remove extra blank columns on the right side of all rows."
    strTemp = strTemp & vbCRLF & "       c. Remove rows that have no data."
    strTemp = strTemp & vbCRLF & "4. Convert CSV files to taget EXCEL file."
    strTemp = strTemp & vbCRLF & " "
    strTemp = strTemp & vbCRLF & "usage:   cleanExcel.vbs <source file> [<target file> [<supress>]] "
    strTemp = strTemp & vbCRLF & "    <source file>  Required    Path with file name to Excel to be processed."
    strTemp = strTemp & vbCRLF & "    <Taget file>   Optional    Name of the new Excel file."
    strTemp = strTemp & vbCRLF & "    <suppress>     Optional    If values set to '1' all dialogs will be suppressed."
    strTemp = strTemp & vbCRLF & " "
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
    If Not(supressNotes=1) Then
        WScript.Echo strTemp
    End If

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
Function IsProcessRunning(strComputer, strProcess)
    Dim Process, strObject
    IsProcessRunning = False
    strObject   = "winmgmts://" & strComputer
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

    If IsProcessRunning(".", "Excel.exe") Then
        bRun=True
        Do While bRun=True
                ' body
            If NOT(IsProcessRunning(".", "Excel.exe")) Then
                bRun=False
            else
                WScript.Sleep(1000)
                sleepCnt=sleepCnt+1
            End If
            If sleepCnt > 100 Then
                WScript.Echo "Excel is running and will not close. Exit VBS"
                WScript.Quit
            End If
        Loop
    End If
    'body
End Function