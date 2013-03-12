VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BM_ErrorLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'This is a simple error logger for use in automated excel reports. It allows code to log errors throughout the report generation process
'then report them onto a new Excel sheet when finished so that end users are notified if any critical errors occured that may result
'in a malformed report. It provides only very basic logging functionality and is intended to augment developer testing.

Private errorLog As String
Private logOrigin As String
Private isInitialized As Boolean

Public Sub Initialize(logName As String)
    logOrigin = logName
    isInitialized = True
End Sub

Public Sub logError(error As String)
    If Not isInitialized Then Err.Raise vbObjectError + 512, "NOD_ErrorLog", "NOD_ErrorLog not initialized."
    errorLog = errorLog & vbNewLine & error
End Sub

Public Sub reportErrors()
    If Not isInitialized Then Err.Raise vbObjectError + 512, "NOD_ErrorLog", "NOD_ErrorLog not initialized."
    If errorLog <> vbNullString Then
        'Delete any outdated sheets
        Dim ws As Worksheet
        For Each ws In ActiveWorkbook.Sheets
            If InStr(1, ws.Name, "Error Log - " & logOrigin) > 0 Then
                Application.DisplayAlerts = False
                ws.Delete
                Application.DisplayAlerts = True
            End If
        Next ws

        Dim errorLogSheet As Worksheet
        Set errorLogSheet = ActiveWorkbook.Sheets.Add(ActiveWorkbook.Sheets(1))
        errorLogSheet.Name = "Error Log - " & logOrigin
        errorLogSheet.Tab.ThemeColor = xlThemeColorAccent2

        With errorLogSheet
            .Cells(2, 2).Value = "Error Log - " & logOrigin
            .Cells(2, 2).Font.Size = 24
            .Cells(2, 2).Borders(xlEdgeBottom).Weight = xlThin
            .Cells(2, 2).HorizontalAlignment = xlCenter
            .Cells(2, 2).VerticalAlignment = xlCenter

            Dim currentRow As Long
            Dim errorLogArray As Variant
            Dim error As Variant
            errorLogArray = Split(errorLog, vbNewLine)
            currentRow = 3

            For Each error In errorLogArray
                .Cells(currentRow, 2).Value = error
                .Cells(currentRow, 2).InsertIndent 1
                currentRow = currentRow + 1
            Next error

            .Columns.AutoFit
            ActiveWindow.DisplayGridlines = False
        End With
    End If
End Sub

