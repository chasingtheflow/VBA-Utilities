Attribute VB_Name = "utils"
Option Explicit 'enforces variable definitions by the compiler

' This file contains a set of Excel VBA Utilities collected by Billy Matthews (billy.matthews@bucknell.edu) over years of Excel centered development.

' The purpose of this module is to provide a centralized location for commonly used generic functions in order to expidite development
' on new projects by providing a basis upon which to bootstrap new projects. Many of these functions were written by me, some have been collected from around the web.

' Some functions assume a link to the "Microsoft Scripting Runtime" module, which is not available on Macs. Though replacement modules can be found online.

Public Function SQLsqueezer(sqlString As String, serverAddress As String, Optional userID As String = vbNullString, Optional password As String = vbNullString, Optional database As String = vbNullString, Optional shouldReturnResultsAtRange As Range = Nothing, Optional shouldReturnListArray As Boolean = True, Optional shouldReturnHeadersAtTopOfRange As Boolean = True) As Variant
    'Executes the given SQL string on the specified database using the specified credentials and can return the results in a variaty of configurations.
    
    'Will return results to an array if specified using a parameter option or will return results to a range on sheet, or any combination
    'Ex: SQLsqueezer sqlString, "server-address", "my_user_id_string", "my_user_password", "database_to_connect_to", shouldReturnResultsAtRange:=ActiveSheet.Range("A1"), shouldReturnListArray:=False, shouldReturnHeadersAtTopOfRange:=True

    Dim adoConnection As Object    'New ADODB.Connection
    Dim adoRcdSource As Object     'New ADODB.Recordset

    Set adoConnection = CreateObject("ADODB.Connection")
    Set adoRcdSource = CreateObject("ADODB.Recordset")
    On Error GoTo Errs:

    'For Excel DB
    adoConnection.Open "Provider=SQLOLEDB.1;Server=" & serverAddress & ";Database=" & database & ";User Id=" & userID & ";Password=" & password & ";"

    If UCase(Left(sqlString, 6)) = "SELECT" Then
        adoRcdSource.Open sqlString, adoConnection, 3
        If shouldReturnListArray = True Then
            If (adoRcdSource.BOF Or adoRcdSource.EOF) = False Then
                SQLsqueezer = adoRcdSource.GetRows
            End If
        End If

        If Not shouldReturnResultsAtRange Is Nothing Then
            With shouldReturnResultsAtRange

                If shouldReturnHeadersAtTopOfRange Then
                    Dim header As Variant, currentColumn As Long
                    currentColumn = 1

                    For Each header In adoRcdSource.Fields
                        .Cells(1, currentColumn).Value = header.Name
                        .Cells(1, currentColumn).Font.Bold = True
                        currentColumn = currentColumn + 1
                    Next header

                    .Cells(2).CopyFromRecordset adoRcdSource
                    freezePanesOnSheet shouldReturnResultsAtRange.Parent, .Range("A2")
                    .Parent.Columns.AutoFit
                Else
                    .Cells(1).CopyFromRecordset adoRcdSource
                End If
            End With
        End If
    Else
        adoConnection.Execute sqlString
    End If

    GoTo NormalExit
Errs:
    MsgBox Err.Description, vbCritical, "Error!"
    Err.Clear: On Error GoTo 0: On Error GoTo -1

NormalExit:
    On Error Resume Next
    adoRcdSource.Close
    On Error GoTo 0
    Set adoConnection = Nothing
    Set adoRcdSource = Nothing
End Function

Public Function getQuarter(Optional offset As Long = 0, Optional ByVal dateSeed As Date = vbNull, Optional errorLog As BM_ErrorLog = Nothing) As String
    'Returns the Fiscal Quarter as a string in the format "Q1-13" where the portion after the dash is the fiscal year. This assumes a fiscal year ending in September.
    'By passing an offset it is possible to request previous or future quarters from the given date. If no date is given then "now" is assumed. This function also works
    'in tandem with a custom error logger also available on GitHub. If you do not want to use the error logger then just remove that portion.
    
    If dateSeed = vbNull Then dateSeed = Now
    dateSeed = DateAdd("q", offset, dateSeed)    'dateSeed + offset * 90
    Select Case Month(dateSeed) Mod 12
        Case 10, 11, 12, 0    'Zero case handles 12 mod 12
            getQuarter = "Q1-" & Right(Year(dateSeed) + 1, 2)
        Case 1, 2, 3
            getQuarter = "Q2-" & Right(Year(dateSeed), 2)
        Case 4, 5, 6
            getQuarter = "Q3-" & Right(Year(dateSeed), 2)
        Case 7, 8, 9
            getQuarter = "Q4-" & Right(Year(dateSeed), 2)
        Case Else
            getQuarter = "INVALID MONTH"
            If Not errorLog Is Nothing Then errorLog.logError "WARNING: Tried to get quarter for " & dateSeed & ", but the date is invalid."
    End Select
End Function

Public Function getLastCompleteQuarter(Optional ByVal dateSeed As Date = vbNull) As Long
    'Returns the last completed quarter as an integer (long) based on a given date. If no date is given then "now" is assumed.
    'Useful in situations where if a quarter is complete you may want to show "Actuals" but if the quarter is incomplete then
    'you may want to show "Forecast" numbers.
    
    If dateSeed = vbNull Then dateSeed = Now
    Select Case Month(dateSeed)
        Case 10, 11, 12
            getLastCompleteQuarter = 0
        Case 1, 2, 3
            getLastCompleteQuarter = 1
        Case 4, 5, 6
            getLastCompleteQuarter = 2
        Case 7, 8, 9
            getLastCompleteQuarter = 3
        Case Else
            getLastCompleteQuarter = -1
    End Select
End Function

Public Function getFirstMonthOfQuarter(Optional ByVal dateSeed As Date = vbNull) As Long
    'Returns the first month of a quarter as a long based on a given date. If no date is given then "now" is assumed.
    'Useful if you want to iterate of each month of a quarter.
    
    If dateSeed = vbNull Then dateSeed = Now
    Select Case Month(dateSeed)
        Case 10, 11, 12
            getFirstMonthOfQuarter = 10
        Case 1, 2, 3
            getFirstMonthOfQuarter = 1
        Case 4, 5, 6
            getFirstMonthOfQuarter = 4
        Case 7, 8, 9
            getFirstMonthOfQuarter = 7
        Case Else
            getFirstMonthOfQuarter = -1
    End Select
End Function

Public Function getFiscalYear(Optional offset As Long = 0, Optional ByVal dateSeed As Date = vbNull) As String
    'Returns the Fiscal Year as a string in the format "FY13". Can be given an offset. If no date is given then "now" is assumed.
    
    If dateSeed = vbNull Then dateSeed = Now
    dateSeed = DateAdd("yyyy", offset, dateSeed)

    Select Case Month(dateSeed)
        Case 10, 11, 12
            getFiscalYear = "FY" & Right(Year(dateSeed) + 1, 2)
        Case Else
            getFiscalYear = "FY" & Right(Year(dateSeed), 2)
    End Select
End Function

Public Sub InitializeColumnHeadersFor(sheetToInitialize As Worksheet, outputDictionary As Dictionary, Optional ByVal headerRow As Long = 1)
    'Parses a header row into a dictionary object for easy access. Repeat headings will contain an appended digit representing which repeat it is.
    'Should consider reworking this as a function rather than a sub for clearer semantics -- however by passing the outputDictionary to the Sub
    'we can initialize it for the user, which is handy.
    
    Dim lastDataColumn As Long
    Dim currentColumn As Long
    Dim currentKey As String
    Dim numberOfRepeats As Long

    Set outputDictionary = New Scripting.Dictionary

    lastDataColumn = sheetToInitialize.UsedRange.Columns.Count

    For currentColumn = 1 To lastDataColumn
        currentKey = Trim(sheetToInitialize.Cells(headerRow, currentColumn).Value)
        currentKey = Trim(currentKey)
        numberOfRepeats = 1
        Do While outputDictionary.exists(currentKey)
            numberOfRepeats = numberOfRepeats + 1
            currentKey = currentKey & " " & numberOfRepeats
        Loop
        If currentKey <> vbNullString Then outputDictionary.Add Key:=currentKey, Item:=currentColumn
    Next currentColumn
End Sub

Public Function createPivotTableOnSheet(ws As Worksheet, dataSheet As Worksheet, Optional atCell As String = "A1") As PivotTable
    'Convenience function for creating a pivot table from a selected data sheet at a specified location
    
    Dim pt As PivotTable
    Set createPivotTableOnSheet = ActiveWorkbook.PivotCaches.Create(xlDatabase, dataSheet.UsedRange).CreatePivotTable(ws.Range(atCell), ws.Name)
End Function

Public Function SheetExists(ByVal sheetName As String) As Boolean
    'Convenience function for checking if a given sheet exists.
    
    On Error Resume Next
    SheetExists = (Sheets(sheetName).Name <> "")
    On Error GoTo 0
End Function

Public Function stripDateFromSheetName(thisSheet As Worksheet) As String
    'Helper function for striping dates from sheets named in the format "SheetName 2-18" where "2-18" represents a date.
    'For reports generated automatically from code, I usually follow the convention of appending the date the sheet was
    'created in the format mm-dd for convenience. This method assists in working with this style of sheet naming.
    
    If inString(thisSheet.Name, "-") Then
        stripDateFromSheetName = Strings.Left(thisSheet.Name, Strings.InStr(1, thisSheet.Name, "-") - 2)
    Else
        stripDateFromSheetName = thisSheet.Name
    End If
End Function

Public Sub freezePanesOnSheet(sheetToFreeze As Worksheet, atPosition As Range)
    'Helper function to freeze panes on a sheet
    'Leaves handling of screen updating to the user
    Dim currentlyActiveSheet As Worksheet
    Set currentlyActiveSheet = ActiveSheet
    
    sheetToFreeze.Activate
    Application.GoTo sheetToFreeze.Range("A1"), True
    sheetToFreeze.Range(atPosition.Address).Select
    ActiveWindow.FreezePanes = True
    
    currentlyActiveSheet.Activate 'reactivate whatever sheet was previously active
End Sub


Public Function inString(stringToSearch As String, stringToLookFor As String, Optional startingAt As Long = 1, Optional compareMethod As VbCompareMethod = vbBinaryCompare) As Boolean
    'Helper function for checking if a string is within another string and returns a boolean.
    
    If InStr(startingAt, stringToSearch, stringToLookFor, compareMethod) > 0 Then
        inString = True
    Else
        inString = False
    End If
End Function

Public Function getUsedRangeOnSheet(Optional thisSheet As Worksheet = Nothing) As Range
    'Returns the used range on a sheet by utilizing the two helper functions below.
    'I've found Excel's builtin "usedRange" function to be unreliable at times. This is a highly usable replacement for data-oriented sheets
    '(such as database dumps) where the data begins in cell A1. Beware that if Row 1 or Column A are empty then this will not behave as you'd expect.
    
    If thisSheet Is Nothing Then Set thisSheet = ActiveSheet
    With thisSheet
        Set getUsedRangeOnSheet = .Range(.Cells(1, 1), .Cells(getLastUsedRowOnSheet(thisSheet), getLastUsedColumnOnSheet(thisSheet)))
    End With
End Function

Public Function getLastUsedRowOnSheet(Optional thisSheet As Worksheet = Nothing) As Long
    'Returns the last used row on a sheet as a long by searching backwards from A1
    If thisSheet Is Nothing Then Set thisSheet = ActiveSheet
    With thisSheet
        getLastUsedRowOnSheet = .Cells.Find("*", [A1], searchorder:=xlByRows, searchdirection:=xlPrevious).Row
    End With
End Function

Public Function getLastUsedColumnOnSheet(Optional thisSheet As Worksheet = Nothing) As Long
    'Returns the last used column on a sheet as a long by searching backward from A1
    If thisSheet Is Nothing Then Set thisSheet = ActiveSheet
    With thisSheet
        getLastUsedColumnOnSheet = .Cells.Find("*", [A1], searchorder:=xlByColumns, searchdirection:=xlPrevious).Column
    End With
End Function

Public Sub addConditionalFormattingForUndefinedOnSheet(sheetToFormat As Worksheet)
' Will colorize all cells in the range that contain the text "Undefined", useful in select scenarios.

    With sheetToFormat.UsedRange
        .FormatConditions.Add xlCellValue, xlEqual, "=""Undefined"""
        .FormatConditions(1).Font.ThemeColor = xlThemeColorAccent2
        .FormatConditions(1).Font.TintAndShade = -0.249946592608417
    End With
End Sub

Public Function getStartOfYearOffset(Optional dateSeed As Date = vbNull) As Long
    'For handling an offset between actual and fiscal years in some scenarios.
    
    If dateSeed = vbNull Then dateSeed = Now
    Select Case Month(dateSeed)
        Case 10, 11, 12
            getStartOfYearOffset = 0
        Case Else
            getStartOfYearOffset = -1
    End Select
End Function

Public Function getActualsOfYearOffset(Optional dateSeed As Date = vbNull) As Long
    'For handling an offset between actual and fiscal years in some scenarios.
    
    If dateSeed = vbNull Then dateSeed = Now
    Select Case Month(dateSeed)
        Case 10, 11, 12
            getActualsOfYearOffset = 1
        Case Else
            getActualsOfYearOffset = 0
    End Select
End Function

Private Function promptForMultipleTextInputs() As Variant
    'Returns an array of paths to .txt files for processing.
    'Allows for batch importing of txt files into excel. Files can then be opened and processed
    'easily using a format like "For Each txtFileToProcess In listOfFilesToProcess". Can be
    'used in tandom with openPipeSeparatedUTF8() to open the files (though you may need
    'to specify different formating/delimiters to meet your needs).
    
    Dim filter As String, title As String
    
    filter = "Text Files (*.txt),*.txt"
    title = "Select multiple txt files to process..."
    
    With Application
        promptForMultipleTextInputs = .GetOpenFilename(filter, 1, title, , True)
    End With
End Function

Private Function openPipeSeparatedUTF8() As Workbook
'Opens a pipe-separated text file, enforcing UTF8 encoding and US English number separators
'Returns workbook object representing processed pipe-separated file
    Dim fn As String

    On Error Resume Next
    fn = Excel.Application.GetOpenFilename( _
         fileFilter:="Text Files (*.txt), *.txt,All Files (*.*),*.*", _
         title:="Open Pipe-Separated Report...")
    If fn <> "False" Then
        Excel.Workbooks.OpenText fileName:=fn, Origin:=msoEncodingUTF8, _
                                 DataType:=xlDelimited, TextQualifier:=xlTextQualifierNone, _
                                 ConsecutiveDelimiter:=False, Tab:=False, Semicolon:=False, _
                                 Comma:=False, Space:=False, other:=True, OtherChar:="|", _
                                 DecimalSeparator:=".", ThousandsSeparator:=","
        Set openPipeSeparatedUTF8 = Excel.ActiveWorkbook
    End If
End Function

Private Sub resaveWorkbookAsXlsx(thisWorkbook As Workbook)
    'Resaves an opened .txt (or any 3-digit extension) workbook as .xlsx
    
    Dim fileName As String
    fileName = Left(thisWorkbook.FullName, Len(thisWorkbook.FullName) - 3) & "xlsx"
    
    On Error Resume Next
    thisWorkbook.SaveAs fileName, xlOpenXMLWorkbook
    On Error GoTo 0
End Sub

Private Function read2DExceptionList(fileName) As Dictionary
' Reads in txt exception list with syntax "Header: Value" and returns a dictionary of dictionaries
' input syntax example:
' Division: Healthcare
' Division: Embedded

' Useful in processing database dumps for exceptions -- though could be adapted for reading in 2 dimensional data for other uses.

    Dim exceptionsDict As Scripting.Dictionary
    Dim tempDict As Scripting.Dictionary
    Dim fHandle As Integer
    Dim fLine As String, strToKeep As String
    Dim pos As Integer
    Dim delimiter As String
    Dim headerMatch As String, valueMatch As String
    Dim errorLog As String
    Dim LineNum As Long

    On Error Resume Next

    delimiter = ":"
    errorLog = ""

    fHandle = FreeFile()
    Open fileName For Input As fHandle
    LineNum = 0

    Set exceptionsDict = New Scripting.Dictionary

    Do While (Not (EOF(fHandle)))
        Line Input #fHandle, fLine
        LineNum = LineNum + 1
        fLine = Trim(fLine)
        If fLine <> "" And Strings.Left(fLine, 1) <> "'" And Strings.Left(fLine, 1) <> "#" Then    'comments delimited by ' or #
            pos = InStr(1, fLine, "'")
            If pos = 0 Then pos = InStr(1, fLine, "#")
            If pos = 0 Then
                strToKeep = fLine
            Else
                strToKeep = Trim(Left(fLine, pos - 1))
            End If

            'split line into header and value:
            pos = Strings.InStr(1, strToKeep, delimiter)
            If pos = 0 Then
                errorLog = errorLog & Chr(9) & _
                           "Missing ':' separator in line " & LineNum & ":   '" & strToKeep & "'" & Chr(13)
            ElseIf pos = 1 Then
                errorLog = errorLog & Chr(9) & _
                           "Column header empty in line " & LineNum & ":   '" & strToKeep & "'" & Chr(13)
            Else
                headerMatch = Strings.Trim(Strings.Left(strToKeep, pos - 1))
                valueMatch = Strings.Trim(Strings.Mid(strToKeep, pos + 1))

                If Not exceptionsDict.exists(headerMatch) Then
                    Set tempDict = New Scripting.Dictionary
                    tempDict.Add Key:=valueMatch, Item:=headerMatch
                    exceptionsDict.Add Key:=headerMatch, Item:=tempDict
                    Set tempDict = Nothing
                Else
                    exceptionsDict(headerMatch).Add Key:=valueMatch, Item:=headerMatch
                End If
            End If
        End If
    Loop
    Close fHandle

    If errorLog <> "" Then
        Dim resp As Integer
        resp = MsgBox("Errors were found in " & fileName & ":" & Chr(13) & Chr(13) & errorLog & _
                      Chr(13) & "Continue anyway?", vbCritical + vbYesNo + vbDefaultButton2, "Error(s) in exception list!")
        If resp = vbNo Then
            Exit Function
        End If
    End If

    Set read2DExceptionList = exceptionsDict
End Function


Function FolderExists(strPath) As Boolean
'Checks for the existance of a folder referenced by strPath
    If Len(Dir(strPath, vbDirectory)) = 0 Then
        FolderExists = False
    Else
        FolderExists = True
    End If
End Function


Function FileExists(FileName As String) As Boolean
'Checks for the existance of a file referenced by FileName
    FileExists = (Dir(FileName) > "")
End Function


Function ReadTextFile(Fname As String, Length As Integer) As Variant
'Reads Length bytes of content from file Fname, and returns the result as a Variant.
    If FileExists(Fname) Then
        Close #1
        
        Open Fname For Input As #1
        ReadTextFile = Input(Length, 1)
        Close 1
    Else
        ReadTextFile = False
    End If

End Function
