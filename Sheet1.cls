VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Const TARGET_FILE_NAME As String = "C:\Users\kazuo\OneDrive\ドキュメント\Excelマクロ\test\test_CelSerchAndMerge.xlsx"
Const TARGET_SHEET_NAME As String = "Target"
Const START_ROW As Long = 3
Const KEY_COL As Long = 1
Const OUTPUT_COL As Long = 2
Const SAERCH_KEY_RANGE As String = "C:C"
Const MERGE_DATA_COL_1 As Long = 1
Const MERGE_DATA_COL_2 As Long = 2


Public Sub CelSerchAndMerge()

    Dim thisWS As Worksheet
    Set thisWS = ThisWorkbook.ActiveSheet
    
    Application.ScreenUpdating = False
    Dim beforeCalculation As XlCalculation
    beforeCalculation = Application.Calculation
    Application.Calculation = xlCalculationManual
    
    Dim targetWB As Workbook
    Set targetWB = Workbooks.Open(TARGET_FILE_NAME, ReadOnly:=True)
    Dim targeWS As Worksheet
    Set targeWS = targetWB.Sheets(TARGET_SHEET_NAME)
        
    Dim tgt As Range  ' 検索するセル範囲
    Dim rng As Range  ' 見つかったRange
    Dim adr As String  ' 最初に見つかったRangeのAddress
    Dim mergeString As String
    
    Dim row As Long
    row = START_ROW
    Do While (thisWS.Cells(row, KEY_COL) <> "")
        
        Set tgt = targeWS.Range(SAERCH_KEY_RANGE)
        Set rng = tgt.Find(What:=thisWS.Cells(row, KEY_COL), LookAt:=xlPart, LookIn:=xlValues)
            
        If rng Is Nothing Then
        Else
            adr = rng.Address
            mergeString = targeWS.Cells(rng.row, MERGE_DATA_COL_1) & vbCrLf & targeWS.Cells(rng.row, MERGE_DATA_COL_2)
            
            Do
                Set rng = tgt.FindNext(After:=rng)
                If rng.Address = adr Then
                    Exit Do
                Else
                    mergeString = mergeString & vbCrLf & targeWS.Cells(rng.row, MERGE_DATA_COL_1) & vbCrLf & targeWS.Cells(rng.row, MERGE_DATA_COL_2)
                End If
            Loop
            thisWS.Cells(row, OUTPUT_COL) = mergeString
        End If
        
        row = row + 1
    Loop
    
    targetWB.Close
    
    Application.ScreenUpdating = True
    Application.Calculation = beforeCalculation


End Sub
