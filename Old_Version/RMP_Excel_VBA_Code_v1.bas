Option Explicit

' --- CONFIGURATION ---
Private Const LOG_PATH As String = "Log\"
Private Const OUT_PATH As String = "Output\"
Private Const DEBUG_PRINT As Boolean = True

' --- GLOBALS (Engine) ---
Private g_FSO As Object
Private g_LogStream As Object
Private g_HeaderMap As Object
Private g_ProcessID As String
Public g_LookupData As Object
Public g_SourceWB As Workbook
Public g_SourceWS As Worksheet
Public g_colIndexDict As Object
Private g_patternDict As Object
Private g_RowIndex As Object

' ============MAIN EXECUTION==================================
Public Sub Calculation()
    Dim tTotal As Double: tTotal = Timer
    Dim bSuccess As Boolean
    
    On Error GoTo MainErr
    ToggleOptimization True
    InitializeGlobals

    Log "========== PROCESS STARTED =========="
    
    If LoadSourceData() Then
        Dim lr As Long: lr = GetLastRow(g_SourceWS)
        Log "Source has " & lr & " rows (including header)"
        
        If lr > 1 Then
            ProcessBatches
            bSuccess = True
            Log "========== COMPLETED | Total: " & Format(Timer - tTotal, "0.00") & "s =========="
        Else
            Log "No data rows to process (only header or empty)", "WARNING"
        End If
    Else
        Log "FAILED to load source data", "ERROR"
    End If

SafeExit:
    CleanUpResources
    ToggleOptimization False
    MsgBox IIf(bSuccess, "Process completed successfully!", "Process failed - check log"), _
           IIf(bSuccess, vbInformation, vbCritical), "Process Complete"
    Exit Sub

MainErr:
    Log "FATAL ERROR: " & Err.Description & " (Err# " & Err.Number & ")", "ERROR"
    Resume SafeExit
End Sub

' =======LOAD SOURCE DATA - UNIVERSAL (CSV + EXCEL)=================
Private Function LoadSourceData() As Boolean
    On Error GoTo LoadErr
    
    LoadSourceData = False
    
    ' Prompt for source file
    Dim fPath As String: fPath = PromptForSourceFile()
    If fPath = "" Then
        Log "No file selected", "WARNING"
        Exit Function
    End If
    
    Log "Loading: " & fPath
    
    ' Validate file exists
    If Not ValidateFileExists(fPath) Then
        Log "File not found: " & fPath, "ERROR"
        Exit Function
    End If
    
    ' Detect file type
    Dim fileExt As String: fileExt = DetectSourceFileType(fPath)
    Log "File type detected: " & UCase(fileExt)
    
    ' Open source workbook
    If Not OpenSourceWorkbook(fPath, fileExt) Then
        Exit Function
    End If
    
    ' Validate workbook opened
    If g_SourceWB Is Nothing Then
        Log "Workbook object is Nothing", "ERROR"
        Exit Function
    End If
    
    ' Validate source worksheet
    If Not ValidateSourceWorksheet() Then
        Exit Function
    End If
    
    ' Get dimensions
    Dim lr As Long: lr = GetLastRow(g_SourceWS)
    Dim lc As Long: lc = GetLastColumn(g_SourceWS)
    
    Log "Loaded: " & lr & " rows ?" & lc & " columns"
    
    ' Validate minimum structure
    If lr < 1 Then
        Log "No rows found", "ERROR"
        g_SourceWB.Close False
        Set g_SourceWS = Nothing
        Set g_SourceWB = Nothing
        Exit Function
    End If
    
    If lc < 1 Then
        Log "No columns found", "ERROR"
        g_SourceWB.Close False
        Set g_SourceWS = Nothing
        Set g_SourceWB = Nothing
        Exit Function
    End If
    
    ' Build header map
    BuildSourceHeaderMap g_SourceWS
    
    LoadSourceData = True
    Exit Function
    
LoadErr:
    Log "LoadSourceData ERROR: " & Err.Description & " (Err# " & Err.Number & ")", "ERROR"
    On Error Resume Next
    If Not g_SourceWB Is Nothing Then g_SourceWB.Close False
    Set g_SourceWS = Nothing
    Set g_SourceWB = Nothing
    LoadSourceData = False
End Function

' Prompt for source file using file dialog
Private Function PromptForSourceFile() As String
    PromptForSourceFile = BrowseForFile()
End Function

' Validate that the specified file exists
Private Function ValidateFileExists(fPath As String) As Boolean
    ValidateFileExists = g_FSO.FileExists(fPath)
End Function

' Detect and normalize the source file type
Private Function DetectSourceFileType(fPath As String) As String
    DetectSourceFileType = LCase(g_FSO.GetExtensionName(fPath))
End Function

' Open source workbook based on file type
Private Function OpenSourceWorkbook(fPath As String, fileExt As String) As Boolean
    On Error GoTo OpenErr
    
    OpenSourceWorkbook = False
    
    Select Case fileExt
        Case "csv"
            ' CSV: Open with Text import
            Set g_SourceWB = Workbooks.Open(FileName:=fPath, _
                                             ReadOnly:=True, _
                                             Local:=True)
        Case "xlsx", "xlsm", "xls"
            ' Excel: Standard open
            Set g_SourceWB = Workbooks.Open(FileName:=fPath, _
                                             ReadOnly:=True, _
                                             UpdateLinks:=0)
        Case Else
            Log "Unsupported file type: " & fileExt, "ERROR"
            Exit Function
    End Select
    
    ' Check if workbook opened successfully
    If Not g_SourceWB Is Nothing Then
        OpenSourceWorkbook = True
    Else
        Log "Failed to open workbook", "ERROR"
    End If
    
    Exit Function
    
OpenErr:
    Log "Failed to open " & UCase(fileExt) & ": " & Err.Description, "ERROR"
    Set g_SourceWB = Nothing
End Function

' Validate source worksheet and assign first worksheet to g_SourceWS
Private Function ValidateSourceWorksheet() As Boolean
    On Error GoTo ValidateErr
    
    ValidateSourceWorksheet = False
    
    ' Get first worksheet
    If g_SourceWB.Sheets.count = 0 Then
        Log "Workbook has no sheets", "ERROR"
        g_SourceWB.Close False
        Set g_SourceWB = Nothing
        Exit Function
    End If
    
    Set g_SourceWS = g_SourceWB.Sheets(1)
    
    If g_SourceWS Is Nothing Then
        Log "Failed to access worksheet", "ERROR"
        g_SourceWB.Close False
        Set g_SourceWB = Nothing
        Set g_SourceWS = Nothing
        Exit Function
    End If
    
    ValidateSourceWorksheet = True
    Exit Function
    
ValidateErr:
    Log "ValidateSourceWorksheet ERROR: " & Err.Description, "ERROR"
    If Not g_SourceWB Is Nothing Then
        g_SourceWB.Close False
        Set g_SourceWB = Nothing
    End If
    Set g_SourceWS = Nothing
End Function

' Build source header map from source worksheet
Private Sub BuildSourceHeaderMap(ws As Worksheet)
    RefreshHeaderMap ws
End Sub

' ==============================================================================
' ROBUST ROW/COLUMN DETECTION
' ==============================================================================
' ===========GetLastRow - ENHANCED FOR CSV COMPATIBILITY=============
Public Function GetLastRow(ws As Worksheet) As Long
    On Error Resume Next
    
    If ws Is Nothing Then
        GetLastRow = 0
        Exit Function
    End If
    
    ' Method 1: Column A (standard approach)
    Dim lr1 As Long: lr1 = ws.Cells(ws.rows.count, 1).End(xlUp).row
    
    ' Method 2: UsedRange (more reliable for CSV and modified files)
    Dim lr2 As Long: lr2 = 0
    If Not ws.UsedRange Is Nothing Then
        lr2 = ws.UsedRange.row + ws.UsedRange.rows.count - 1
    End If
    
    ' Method 3: Direct UsedRange.Rows.count
    Dim lr3 As Long: lr3 = ws.UsedRange.rows.count
    
    ' Method 4: Special handling for CSV - check if we have data beyond header
    Dim lr4 As Long: lr4 = 0
    If lr1 <= 1 And lr3 > 1 Then
        ' CSV edge case: Method 1 fails but UsedRange shows data
        lr4 = lr3
    End If
    
    ' Use maximum of all methods
    GetLastRow = Application.Max(lr1, lr2, lr3, lr4)
    
    ' Final validation for CSV files
    If GetLastRow < 1 Then
        ' Last resort: Count non-empty cells in column A starting from row 2
        Dim cell As Range
        Set cell = ws.Cells(2, 1)
        If Not IsEmpty(cell) And Not IsError(cell) Then
            GetLastRow = ws.Cells(ws.rows.count, 1).End(xlUp).row
            If GetLastRow < 1 Then GetLastRow = ws.UsedRange.rows.count
        End If
    End If
    
    ' Ensure minimum 1 if data exists
    If GetLastRow < 1 Then
        If ws.UsedRange.rows.count > 0 Then
            GetLastRow = ws.UsedRange.rows.count
        Else
            GetLastRow = 1
        End If
    End If
    
    On Error GoTo 0
End Function


' ===========GetLastColumn - ROBUST DETECTION=============
Private Function GetLastColumn(ws As Worksheet) As Long
    On Error Resume Next
    
    If ws Is Nothing Then
        GetLastColumn = 0
        Exit Function
    End If
    
    ' Method 1: Row 1
    Dim lc1 As Long: lc1 = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
    
    ' Method 2: UsedRange
    Dim lc2 As Long
    If Not ws.UsedRange Is Nothing Then
        lc2 = ws.UsedRange.Columns.count + ws.UsedRange.Column - 1
    End If
    
    ' Use maximum
    GetLastColumn = Application.Max(lc1, lc2)
    
    ' Ensure minimum 1
    If GetLastColumn < 1 Then GetLastColumn = 1
    
    On Error GoTo 0
End Function


' ==============================================================================
' PROCESS BATCHES - OPTIMIZED
' ==============================================================================
Private Sub ProcessBatches()
    On Error GoTo ProcessErr
    
    ' Get distinct MI_NOs
    Dim colUnique As Collection: Set colUnique = GetDistinctMI_NOs()
    If colUnique Is Nothing Or colUnique.count = 0 Then
        Log "No distinct MI_NOs found", "WARNING"
        Exit Sub
    End If
    
    Log "Found " & colUnique.count & " distinct MI_NOs"
    
    ' Build row index for fast filtering
    BuildMI_NORowIndex()
    
    ' Calculate optimal batch size
    Dim batchSize As Long: batchSize = CalculateOptimalBatchSize(colUnique.count)
    Dim totalBatches As Long: totalBatches = CalculateTotalBatches(colUnique.count, batchSize)
    Log "Batch size: " & batchSize & " | Total batches: " & totalBatches
    
    ' Execute batch loop
    ExecuteBatchLoop colUnique, batchSize, totalBatches
    
    Log "All " & totalBatches & " batches processed"
    Exit Sub
    
ProcessErr:
    Log "ProcessBatches ERROR: " & Err.Description, "ERROR"
    Err.Raise Err.Number, "ProcessBatches", Err.Description
End Sub

' Get collection of unique MI_NOs
Private Function GetDistinctMI_NOs() As Collection
    Set GetDistinctMI_NOs = GetDistinctIDs(g_SourceWS)
End Function

' Build g_RowIndex for fast filtering
Private Sub BuildMI_NORowIndex()
    Log "Building row index..."
    Dim tIndex As Double: tIndex = Timer
    BuildRowIndex g_SourceWS
    Log "Row index built in " & Format(Timer - tIndex, "0.00") & "s"
End Sub

' Calculate optimal batch size based on total distinct IDs
Private Function CalculateOptimalBatchSize(totalIDs As Long) As Long
    CalculateOptimalBatchSize = CalculateBatchSize(totalIDs)
End Function

' Calculate total number of batches
Private Function CalculateTotalBatches(totalIDs As Long, batchSize As Long) As Long
    CalculateTotalBatches = -Int(-totalIDs / batchSize)
End Function

' Loop through batches and trigger ProcessSingleBatch
Private Sub ExecuteBatchLoop(colUnique As Collection, batchSize As Long, totalBatches As Long)
    On Error GoTo BatchLoopErr
    
    Dim batchIdx As Long, i As Long, k As Long, endIdx As Long
    
    ' Process each batch
    For i = 1 To colUnique.count Step batchSize
        Dim tBatch As Double: tBatch = Timer
        batchIdx = batchIdx + 1
        
        Log "========== Batch " & batchIdx & "/" & totalBatches & " =========="
        
        ' Build batch dictionary
        Dim batchIDs As Object: Set batchIDs = CreateObject("Scripting.Dictionary")
        On Error Resume Next
        endIdx = Application.Min(i + batchSize - 1, colUnique.count)
        For k = i To endIdx
            batchIDs(colUnique(k)) = Empty
        Next k
        On Error GoTo BatchLoopErr
        
        If batchIDs.count = 0 Then
            Log "Batch " & batchIdx & " has 0 IDs - SKIPPED", "WARNING"
            GoTo NextBatch
        End If
        
        Log "  Contains " & batchIDs.count & " MI_NOs (IDs " & i & " to " & endIdx & ")"
        
        ' Process batch
        If ProcessSingleBatch(batchIDs, batchIdx) Then
            Log "  COMPLETED in " & Format(Timer - tBatch, "0.00") & "s"
        Else
            Log "  FAILED", "ERROR"
        End If
        
NextBatch:
        Set batchIDs = Nothing
        If batchIdx Mod 5 = 0 Then DoEvents
    Next i
    
    Exit Sub
    
BatchLoopErr:
    Log "ExecuteBatchLoop ERROR at Batch " & batchIdx & ": " & Err.Description, "ERROR"
    Err.Raise Err.Number, "ExecuteBatchLoop", Err.Description
End Sub

' ===========ProcessSingleBatch=============
Private Function ProcessSingleBatch(batchIDs As Object, bIdx As Long) As Boolean
    On Error GoTo BatchErr
    
    ' Create temporary workbook
    Dim wbTemp As Workbook, wsTemp As Worksheet
    CreateTemporaryWorkbook wbTemp, wsTemp
    If wbTemp Is Nothing Or wsTemp Is Nothing Then
        Log "    Failed to create temp workbook", "ERROR"
        ProcessSingleBatch = False
        Exit Function
    End If
    
    ' Filter data for current batch
    If Not FilterBatchData(wsTemp, batchIDs) Then
        GoTo CleanBatch
    End If
    
    Dim filteredRows As Long: filteredRows = GetLastRow(wsTemp)
    Log "    Filtered " & (filteredRows - 1) & " data rows"
    
    If filteredRows < 2 Then
        Log "    No matching data - SKIPPED", "WARNING"
        GoTo CleanBatch
    End If
    
    ' Refresh header map for temp worksheet
    RefreshHeaderMap wsTemp
    
    ' Execute batch helpers
    ExecuteBatchHelpers wsTemp
    
    ' Save batch output
    If SaveBatchOutput(wbTemp, bIdx) Then
        ProcessSingleBatch = True
    End If
    
CleanBatch:
    On Error Resume Next
    If Not wbTemp Is Nothing Then wbTemp.Close False
    Set wbTemp = Nothing
    Set wsTemp = Nothing
    On Error GoTo 0
    Exit Function
    
BatchErr:
    Log "BATCH ERROR " & bIdx & ": " & Err.Description, "ERROR"
    Resume CleanBatch
End Function

' Create and return new temp workbook/worksheet
Private Sub CreateTemporaryWorkbook(ByRef wbTemp As Workbook, ByRef wsTemp As Worksheet)
    On Error GoTo CreateErr
    
    Log "    Creating temp workbook..."
    Set wbTemp = Workbooks.Add(xlWBATWorksheet)
    Set wsTemp = wbTemp.Sheets(1)
    
    Exit Sub
    
CreateErr:
    Log "    ERROR creating temp workbook: " & Err.Description, "ERROR"
    Set wbTemp = Nothing
    Set wsTemp = Nothing
End Sub

' Call FilterDataToSheet for current batch
Private Function FilterBatchData(wsTemp As Worksheet, batchIDs As Object) As Boolean
    On Error GoTo FilterErr
    
    FilterBatchData = False
    
    Log "    Filtering data..."
    FilterDataToSheet g_SourceWS, wsTemp, batchIDs
    
    FilterBatchData = True
    Exit Function
    
FilterErr:
    Log "    FILTER ERROR: " & Err.Description, "ERROR"
End Function

' Run applicable helpers on temp sheet
Private Sub ExecuteBatchHelpers(wsTemp As Worksheet)
    On Error GoTo HelperErr
    
    Log "    Running helpers..."
    RunAllHelpers wsTemp
    
    Exit Sub
    
HelperErr:
    Log "    ERROR executing helpers: " & Err.Description, "ERROR"
End Sub

' Save temp workbook to CSV with standardized naming
Private Function SaveBatchOutput(wbTemp As Workbook, bIdx As Long) As Boolean
    On Error GoTo SaveErr
    
    SaveBatchOutput = False
    
    Log "    Saving output..."
    Dim fName As String
    fName = OUT_PATH & g_FSO.GetBaseName(g_SourceWB.name) & _
            "_Batch_" & Format(bIdx, "000") & "_" & g_ProcessID & ".csv"
    
    Application.DisplayAlerts = False
    wbTemp.SaveAs fName, xlCSV
    Application.DisplayAlerts = True
    
    Log "    Saved: " & g_FSO.GetFileName(fName)
    SaveBatchOutput = True
    
    Exit Function
    
SaveErr:
    Log "    ERROR saving output: " & Err.Description, "ERROR"
End Function

' ===========RunAllHelpers=============
Private Sub RunAllHelpers(ws As Worksheet)
    'ExecuteHelper ws, "Helper_01_SubGroupMapping"
    'ExecuteHelper ws, "Helper_02_Source"
    'ExecuteHelper ws, "Helper_03_NB_and_RI"
    'ExecuteHelper ws, "Helper_04_DefaultRate"
    'ExecuteHelper ws, "Helper_05_LTVPropValue"
    'ExecuteHelper ws, "Helper_06_AssumptionValuation"
    'ExecuteHelper ws, "Helper_07_Policies_Date"
    'ExecuteHelper ws, "Helper_08_ScenarioMerge"
    'ExecuteHelper ws, "Helper_09_ReinsurerMapping"
    'ExecuteHelper ws, "Helper_10_AddGMR_MTHLY"
    'ExecuteHelper ws, "Helper_11_AddMaturity_Term"
    
    'ExecuteHelper ws, "Helper_13_Projection_Expand"
    'RefreshHeaderMap ws
    
    'ExecuteHelper ws, "Helper_14_AddRemaining_Term"
    'ExecuteHelper ws, "Helper_15_YrMth_Indicator"
    'ExecuteHelper ws, "Helper_16_DefaultPattern"
    'ExecuteHelper ws, "Helper_17_ClaimRate"
    'ExecuteHelper ws, "Helper_18_CalculateOMCR"
    'ExecuteHelper ws, "Helper_20_OMV_and_PctOMV_Combined"
    'ExecuteHelper ws, "Helper_21_CPR_SMM_Combined"
    'ExecuteHelper ws, "Helper_23_AcquisitionExpense"
    'ExecuteHelper ws, "Helper_24_PolicyForceAndOPB_Combined"
    'RefreshHeaderMap ws
    
    'ExecuteHelper ws, "Helper_24A_Duration_Plus_and_Min"
    'ExecuteHelper ws, "Helper_24B_Next_Policy_In_Force"
    'ExecuteHelper ws, "Helper_24C_PreviousValues"
    'RefreshHeaderMap ws
    
    'ExecuteHelper ws, "Helper_25_FixedAssumedSeverity"
    'ExecuteHelper ws, "Helper_26_InflationFactor"
    'ExecuteHelper ws, "Helper_27_MaintenanceExpense"
    'ExecuteHelper ws, "Helper_28_Commission_and_Commission_Recovery"
    'ExecuteHelper ws, "Helper_29_RI_Policies_Count"

    'ExecuteHelper ws, "Helper_30_RI_Premium"
    'ExecuteHelper ws, "Helper_31_RiskInForce_and_DefaultClaimOutgo"
    'ExecuteHelper ws, "Helper_32_RI_Collateral"
    'ExecuteHelper ws, "Helper_33_RI_NPR"

End Sub

Private Sub ExecuteHelper(ws As Worksheet, procName As String)
    Dim t As Double: t = Timer
    On Error Resume Next
    Application.Run procName, ws
    If Err.Number <> 0 Then
        Log "      " & procName & " ERROR: " & Err.Description, "ERROR"
    Else
        Log "      " & procName & ": " & Format(Timer - t, "0.000") & "s", "SUCCESS"
    End If
    Err.Clear
End Sub

' ==============================================================================
' FILTERING ENGINE
' ==============================================================================
Private Sub BuildRowIndex(ws As Worksheet)
    On Error GoTo IndexErr
    
    Set g_RowIndex = CreateObject("Scripting.Dictionary")
    
    Dim lr As Long: lr = GetLastRow(ws)
    If lr < 2 Then Exit Sub
    
    Dim vMI As Variant
    On Error Resume Next
    vMI = ws.Range(ws.Cells(2, 1), ws.Cells(lr, 1)).Value2
    If Err.Number <> 0 Then
        Log "  BuildRowIndex: Failed to read MI_NO column", "ERROR"
        Exit Sub
    End If
    On Error GoTo IndexErr
    
    vMI = SafeArray2D(vMI)
    Dim rowCount As Long: rowCount = UBound(vMI, 1)
    
    Dim i As Long, miNo As String, rowCol As Collection
    
    For i = 1 To rowCount
        On Error Resume Next
        miNo = Trim(CStr(vMI(i, 1)))
        Err.Clear
        On Error GoTo IndexErr
        
        If Len(miNo) > 0 Then
            If Not g_RowIndex.Exists(miNo) Then
                Set rowCol = New Collection
                g_RowIndex.Add miNo, rowCol
            Else
                Set rowCol = g_RowIndex(miNo)
            End If
            rowCol.Add i + 1
        End If
    Next i
    
    Log "  Indexed " & g_RowIndex.count & " unique MI_NOs"
    
    vMI = Empty
    Set rowCol = Nothing
    Exit Sub
    
IndexErr:
    Log "BuildRowIndex ERROR: " & Err.Description, "ERROR"
    Set g_RowIndex = CreateObject("Scripting.Dictionary")
End Sub

' 复制源工作表的表头到目标工作表
Private Sub CopySourceHeader(srcWS As Worksheet, destWS As Worksheet)
    srcWS.Rows(1).Copy destWS.Rows(1)
End Sub

' 从源工作表读取数据并返回一个变体数组
Private Function ReadSourceDataArray(srcWS As Worksheet, lr As Long, lc As Long) As Variant
    On Error GoTo ReadErr
    
    ReadSourceDataArray = srcWS.Range("A1", srcWS.Cells(lr, lc)).Value2
    Exit Function
    
ReadErr:
    Log "      ERROR reading source: " & Err.Description, "ERROR"
    ReadSourceDataArray = Empty
End Function

' 根据有效ID和全局行索引过滤数据数组
Private Function FilterDataByValidIDs(vSrc As Variant, srcRows As Long, srcCols As Long, validIDs As Object) As Variant
    Dim vRes As Variant: ReDim vRes(1 To srcRows, 1 To srcCols)
    Dim outR As Long, r As Long, c As Long
    Dim miNo As Variant, rowCol As Collection, rowNum As Variant
    
    ' 使用g_RowIndex进行高效查找（如果可用）
    If Not g_RowIndex Is Nothing And g_RowIndex.Count > 0 Then
        For Each miNo In validIDs.Keys
            If g_RowIndex.Exists(CStr(miNo)) Then
                Set rowCol = g_RowIndex(CStr(miNo))
                For Each rowNum In rowCol
                    r = CLng(rowNum)
                    If r >= 2 And r <= srcRows Then
                        outR = outR + 1
                        For c = 1 To srcCols
                            vRes(outR, c) = vSrc(r, c)
                        Next c
                    End If
                Next rowNum
            End If
        Next miNo
    Else
        ' 回退方案：全表扫描
        For r = 2 To srcRows
            On Error Resume Next
            miNo = Trim(CStr(vSrc(r, 1)))
            Err.Clear
            On Error GoTo FilterErr
            
            If Len(CStr(miNo)) > 0 And validIDs.Exists(CStr(miNo)) Then
                outR = outR + 1
                For c = 1 To srcCols
                    vRes(outR, c) = vSrc(r, c)
                Next c
            End If
        Next r
    End If
    
    ' 将结果重新定义为实际大小
    If outR > 0 Then
        Dim vFinal As Variant: ReDim vFinal(1 To outR, 1 To srcCols)
        For r = 1 To outR
            For c = 1 To srcCols
                vFinal(r, c) = vRes(r, c)
            Next c
        Next r
        FilterDataByValidIDs = vFinal
    Else
        FilterDataByValidIDs = Empty
    End If
    
    Exit Function
    
FilterErr:
    Log "FilterDataByValidIDs ERROR: " & Err.Description, "ERROR"
    FilterDataByValidIDs = Empty
End Function

' 将过滤后的数据数组写入目标工作表
Private Sub WriteFilteredData(destWS As Worksheet, vRes As Variant)
    If Not IsEmpty(vRes) Then
        Dim outRows As Long: outRows = UBound(vRes, 1)
        Dim outCols As Long: outCols = UBound(vRes, 2)
        destWS.Range("A2").Resize(outRows, outCols).Value2 = vRes
    End If
End Sub

' 主要的过滤流程，调用上述专用函数
Private Sub FilterDataToSheet(srcWS As Worksheet, destWS As Worksheet, validIDs As Object)
    On Error GoTo FilterErr
    
    If srcWS Is Nothing Then Err.Raise 91, , "Source worksheet is Nothing"
    
    ' 复制表头
    CopySourceHeader srcWS, destWS
    
    Dim lr As Long: lr = GetLastRow(srcWS)
    If lr < 1 Then
        Log "      No data in source", "WARNING"
        Exit Sub
    End If
    
    Dim lc As Long: lc = GetLastColumn(srcWS)
    
    If lr < 2 Then Exit Sub
    
    ' 读取源数据
    Dim vSrc As Variant
    vSrc = ReadSourceDataArray(srcWS, lr, lc)
    If IsEmpty(vSrc) Then Exit Sub
    
    vSrc = SafeArray2D(vSrc)
    Dim srcRows As Long: srcRows = UBound(vSrc, 1)
    Dim srcCols As Long: srcCols = UBound(vSrc, 2)
    
    ' 筛选数据
    Dim vRes As Variant
    vRes = FilterDataByValidIDs(vSrc, srcRows, srcCols, validIDs)
    
    ' 写入目标工作表
    WriteFilteredData destWS, vRes
    
    ' 清理内存
    vSrc = Empty
    vRes = Empty
    Exit Sub
    
FilterErr:
    Log "FilterDataToSheet ERROR: " & Err.Description, "ERROR"
    Err.Raise Err.Number, "FilterDataToSheet", Err.Description
End Sub

Private Function GetDistinctIDs(ws As Worksheet) As Collection
    On Error GoTo IDErr
    
    Set GetDistinctIDs = New Collection
    
    Dim lr As Long: lr = GetLastRow(ws)
    If lr < 2 Then Exit Function
    
    Dim v As Variant
    On Error Resume Next
    v = ws.Range(ws.Cells(2, 1), ws.Cells(lr, 1)).Value2
    If Err.Number <> 0 Then
        Set GetDistinctIDs = New Collection
        Exit Function
    End If
    On Error GoTo IDErr
    
    v = SafeArray2D(v)
    Dim rowCount As Long: rowCount = UBound(v, 1)
    
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Dim i As Long, s As String
    
    For i = 1 To rowCount
        On Error Resume Next
        s = Trim(CStr(v(i, 1)))
        Err.Clear
        On Error GoTo IDErr
        
        If Len(s) > 0 And Not d.Exists(s) Then
            d(s) = Empty
            GetDistinctIDs.Add s
        End If
    Next i
    
    v = Empty
    Set d = Nothing
    Exit Function
    
IDErr:
    Log "GetDistinctIDs ERROR: " & Err.Description, "ERROR"
    Set GetDistinctIDs = New Collection
End Function

' ==============================================================================
' UTILITIES
' ==============================================================================
' ================= ENHANCED AddColumnIfNotExist - HANDLES ALL SITUATIONS =================
' Merges functionality of AddColumnIfNotExist and AddCol into a single robust function
Public Function AddColumn(ws As Worksheet, colName As String) As Long
    On Error GoTo ErrHandler
    
    Dim ucName As String
    ucName = UCase(Trim(Replace(colName, Chr(160), " "))) ' Normalize
    
    ' Fail-fast: Validate input
    If Len(ucName) = 0 Then
        Log "AddColumn: Empty column name provided", "WARNING"
        AddColumn = 0
        Exit Function
    End If
    
    ' Check if column already exists in header map
    If g_HeaderMap.Exists(ucName) Then
        ' Verify the column still exists and matches
        On Error Resume Next
        Dim existingCol As Long: existingCol = g_HeaderMap(ucName)
        Dim existingHeader As String: existingHeader = ws.Cells(1, existingCol).value
        If Err.Number = 0 And UCase(Trim(existingHeader)) = ucName Then
            ' Valid existing column
            AddColumn = existingCol
            Exit Function
        Else
            ' Stale map entry, remove it
            g_HeaderMap.Remove ucName
        End If
        On Error GoTo ErrHandler
    End If
    
    ' Find the next available column using efficient method (from AddCol)
    Dim targetCol As Long
    targetCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
    If targetCol = 0 Then targetCol = 1
    If ws.Cells(1, targetCol).value <> "" Then targetCol = targetCol + 1
    
    ' Validate target column bounds
    If targetCol > 16384 Then
        Log "AddColumn: Cannot add column " & colName & " - Excel column limit reached", "ERROR"
        AddColumn = 0
        Exit Function
    End If
    
    ' Add header to worksheet
    ws.Cells(1, targetCol).value = colName
    
    ' Update header map
    On Error Resume Next
    g_HeaderMap(ucName) = targetCol
    If Err.Number <> 0 Then
        Log "AddColumn: ERROR updating HeaderMap: " & Err.Description, "WARNING"
        Err.Clear
    End If
    On Error GoTo ErrHandler
    
    AddColumn = targetCol
    Exit Function

ErrHandler:
    Log "AddColumn: CRITICAL ERROR for column " & colName & ": " & Err.Description & " (Err#" & Err.Number & ")", "ERROR"
    
    ' Attempt recovery: try direct write to next available column
    On Error Resume Next
    Dim recoveryCol As Long: recoveryCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column + 1
    If recoveryCol > 0 And recoveryCol <= 16384 Then
        ws.Cells(1, recoveryCol).value = colName
        If Err.Number = 0 Then
            g_HeaderMap(ucName) = recoveryCol
            AddColumn = recoveryCol
            Log "AddColumn: Recovery successful - added " & colName & " at column " & recoveryCol, "WARNING"
            Exit Function
        End If
    End If
    
    AddColumn = 0
    Err.Clear
    On Error GoTo 0
End Function

' Keep backward compatibility
Public Sub AddColumnIfNotExist(ws As Worksheet, colName As String)
    AddColumn ws, colName
End Sub

Public Function GetColumnIndex(ws As Worksheet, colName As String) As Long
    On Error Resume Next
    Dim ucName As String: ucName = UCase(Trim(colName))
    If g_HeaderMap.Exists(ucName) Then
        GetColumnIndex = g_HeaderMap(ucName)
    Else
        GetColumnIndex = 0
    End If
    On Error GoTo 0
End Function

Private Sub RefreshHeaderMap(ws As Worksheet)
    On Error GoTo RefreshErr
    
    g_HeaderMap.RemoveAll
    
    Dim vHead As Variant
    On Error Resume Next
    vHead = ws.rows(1).Value2
    If Err.Number <> 0 Then Exit Sub
    On Error GoTo RefreshErr
    
    vHead = SafeArray2D(vHead)
    Dim colCount As Long: colCount = UBound(vHead, 2)
    Dim c As Long, colName As String
    
    For c = 1 To colCount
        On Error Resume Next
        If Not IsError(vHead(1, c)) And Not IsEmpty(vHead(1, c)) Then
            colName = Trim(CStr(vHead(1, c)))
            If Len(colName) > 0 Then g_HeaderMap(UCase(colName)) = c
        End If
        Err.Clear
        On Error GoTo RefreshErr
    Next c
    
    vHead = Empty
    Exit Sub
    
RefreshErr:
    Log "RefreshHeaderMap ERROR: " & Err.Description, "ERROR"
End Sub

' ======== FIXED NZ FUNCTIONS - EXPLICIT PARAMETERS ========
Public Function NzLong(data As Variant, row As Long, col As Long) As Long
    On Error Resume Next
    If IsArray(data) Then
        If row <= UBound(data, 1) And col <= UBound(data, 2) Then
            Dim v As Variant: v = data(row, col)
            If IsEmpty(v) Or IsNull(v) Or IsError(v) Then
                NzLong = 0
            ElseIf IsNumeric(v) Then
                NzLong = CLng(v)
            Else
                NzLong = 0
            End If
        Else
            NzLong = 0
        End If
    Else
        NzLong = 0
    End If
    On Error GoTo 0
End Function

Public Function NzDouble(data As Variant, r As Long, c As Long) As Double
    On Error Resume Next
    Dim v As Variant: v = data(r, c)
    If IsEmpty(v) Or IsNull(v) Or IsError(v) Then
        NzDouble = 0#
    ElseIf IsNumeric(v) Then
        NzDouble = CDbl(v)
        If Err.Number <> 0 Then NzDouble = 0#
    Else
        NzDouble = 0#
    End If
    On Error GoTo 0
End Function

Public Function NzString(data As Variant, r As Long, c As Long) As String
    On Error Resume Next
    Dim v As Variant: v = data(r, c)
    If IsEmpty(v) Or IsNull(v) Or IsError(v) Then
        NzString = ""
    Else
        NzString = Trim(CStr(v))
        If Err.Number <> 0 Then NzString = ""
    End If
    On Error GoTo 0
End Function

' ======== INLINE NZ FUNCTIONS - EXPLICIT PARAMETERS ========
' =================== INLINE NZDOUBLE (OPTIMIZED) ===================
Private Function InlineNzDouble(v As Variant) As Double
    On Error Resume Next
    Select Case varType(v)
        Case 2 To 6, 14, 17, 20: InlineNzDouble = CDbl(v)
        Case 8: InlineNzDouble = CDbl(v)
        Case Else: InlineNzDouble = 0#
    End Select
    If Err.Number <> 0 Then InlineNzDouble = 0#: Err.Clear
End Function

' =================== UNIFIED NULL-SAFE CONVERSION SYSTEM ===================
' Core validation function that checks for empty/null/error values
Private Function NzConvert(ByVal val As Variant) As Variant
    ' Fail-fast: Handle empty/null/error values immediately
    If IsEmpty(val) Or IsNull(val) Or IsError(val) Then
        NzConvert = Empty
        Exit Function
    End If
    
    NzConvert = val
End Function

' Converts a value from an array to a Long with null/error safety
Public Function NzToLong(data As Variant, Optional row As Long = 1, Optional col As Long = 1) As Long
    On Error Resume Next
    
    Dim val As Variant
    If IsArray(data) Then
        If row <= UBound(data, 1) And col <= UBound(data, 2) Then
            val = NzConvert(data(row, col))
            If IsEmpty(val) Then
                NzToLong = 0
                Exit Function
            End If
            
            If IsNumeric(val) Then
                NzToLong = CLng(val)
            Else
                NzToLong = 0
            End If
        Else
            NzToLong = 0
        End If
    Else
        val = NzConvert(data)
        If IsEmpty(val) Then
            NzToLong = 0
            Exit Function
        End If
        
        If IsNumeric(val) Then
            NzToLong = CLng(val)
        Else
            NzToLong = 0
        End If
    End If
    
    If Err.Number <> 0 Then NzToLong = 0
    On Error GoTo 0
End Function

' Converts a value from an array to a Double with null/error safety
Public Function NzToDouble(data As Variant, Optional row As Long = 1, Optional col As Long = 1) As Double
    On Error Resume Next
    
    Dim val As Variant
    If IsArray(data) Then
        If row <= UBound(data, 1) And col <= UBound(data, 2) Then
            val = NzConvert(data(row, col))
            If IsEmpty(val) Then
                NzToDouble = 0#
                Exit Function
            End If
            
            If IsNumeric(val) Then
                NzToDouble = CDbl(val)
            Else
                NzToDouble = 0#
            End If
        Else
            NzToDouble = 0#
        End If
    Else
        val = NzConvert(data)
        If IsEmpty(val) Then
            NzToDouble = 0#
            Exit Function
        End If
        
        If IsNumeric(val) Then
            NzToDouble = CDbl(val)
        Else
            NzToDouble = 0#
        End If
    End If
    
    If Err.Number <> 0 Then NzToDouble = 0#
    On Error GoTo 0
End Function

' Converts a value from an array to a String with null/error safety
Public Function NzToString(data As Variant, Optional row As Long = 1, Optional col As Long = 1) As String
    On Error Resume Next
    
    Dim val As Variant
    If IsArray(data) Then
        If row <= UBound(data, 1) And col <= UBound(data, 2) Then
            val = NzConvert(data(row, col))
            If IsEmpty(val) Then
                NzToString = ""
                Exit Function
            End If
            
            NzToString = Trim(CStr(val))
        Else
            NzToString = ""
        End If
    Else
        val = NzConvert(data)
        If IsEmpty(val) Then
            NzToString = ""
            Exit Function
        End If
        
        NzToString = Trim(CStr(val))
    End If
    
    If Err.Number <> 0 Then NzToString = ""
    On Error GoTo 0
End Function

' Fast inline conversion of a variant to Double with null/error safety
Public Function NzToDoubleFast(val As Variant) As Double
    On Error Resume Next
    
    Dim safeVal As Variant: safeVal = NzConvert(val)
    If IsEmpty(safeVal) Then
        NzToDoubleFast = 0#
        Exit Function
    End If
    
    Select Case VarType(safeVal)
        Case 2 To 6, 14, 17, 20  ' Numeric types
            NzToDoubleFast = CDbl(safeVal)
        Case 8  ' String
            NzToDoubleFast = CDbl(safeVal)
        Case Else
            NzToDoubleFast = 0#
    End Select
    
    If Err.Number <> 0 Then NzToDoubleFast = 0#
    On Error GoTo 0
End Function

Private Sub InitializeGlobalCollections()
    ' Initialize all global collection objects
    Set g_FSO = CreateObject("Scripting.FileSystemObject")
    Set g_HeaderMap = CreateObject("Scripting.Dictionary")
    Set g_LookupData = CreateObject("Scripting.Dictionary")
    Set g_RowIndex = CreateObject("Scripting.Dictionary")
    Set g_colIndexDict = g_HeaderMap
End Sub

Private Sub EnsureRequiredFolders()
    ' Ensure required folder structure exists
    If Not g_FSO.FolderExists(OUT_PATH) Then g_FSO.CreateFolder OUT_PATH
    If Not g_FSO.FolderExists(LOG_PATH) Then g_FSO.CreateFolder LOG_PATH
End Sub

Private Sub InitializeProcessLogging()
    ' Generate process ID and initialize logging
    g_ProcessID = Format(Now, "yyyymmdd_hhmmss")
    
    Dim logFile As String: logFile = LOG_PATH & "Run_" & g_ProcessID & ".txt"
    Set g_LogStream = g_FSO.CreateTextFile(logFile, True)
    
    Log "Process ID: " & g_ProcessID
    Log "Log file: " & logFile
End Sub

Private Sub InitializeGlobals()
    ' Initialize all globals by calling specialized initialization functions
    InitializeGlobalCollections
    EnsureRequiredFolders
    InitializeProcessLogging
End Sub

Private Sub CleanUpResources()
    On Error Resume Next
    If Not g_LogStream Is Nothing Then 
        g_LogStream.Close
        Set g_LogStream = Nothing
    End If
    If Not g_SourceWB Is Nothing Then 
        g_SourceWB.Close False
        Set g_SourceWB = Nothing
    End If
    Set g_SourceWS = Nothing
    Set g_FSO = Nothing
    Set g_HeaderMap = Nothing
    Set g_LookupData = Nothing
    Set g_RowIndex = Nothing
    Set g_colIndexDict = Nothing
    Set g_patternDict = Nothing
    On Error GoTo 0
End Sub

Private Sub ToggleOptimization(bOn As Boolean)
    With Application
        .ScreenUpdating = Not bOn
        .Calculation = IIf(bOn, xlCalculationManual, xlCalculationAutomatic)
        .EnableEvents = Not bOn
        .DisplayAlerts = Not bOn
        .StatusBar = Not bOn
    End With
End Sub

Private Sub Log(msg As String, Optional sType As String = "INFO")
    Dim s As String: s = Format(Now, "hh:mm:ss") & " [" & sType & "] " & msg
    If DEBUG_PRINT Then Debug.Print s
    On Error Resume Next
    If Not g_LogStream Is Nothing Then 
        g_LogStream.WriteLine s
    Else
        ' Fallback to immediate window if log stream unavailable
        Debug.Print s
    End If
    On Error GoTo 0
End Sub

Private Function CalculateBatchSize(total As Long) As Long
    Select Case total
        Case Is > 10000: CalculateBatchSize = 3000
        Case Is > 5000: CalculateBatchSize = 2000
        Case Is > 2000: CalculateBatchSize = 1500
        Case Is > 500: CalculateBatchSize = Application.Min(total, 1000)
        Case Else: CalculateBatchSize = total
    End Select
End Function

Private Function BrowseForFile() As String
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "Select Input File"
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xlsx;*.xls;*.xlsm"
        .Filters.Add "CSV Files", "*.csv"
        .Filters.Add "All Files", "*.*"
        If .Show = -1 Then BrowseForFile = .SelectedItems(1)
    End With
End Function


' ==============================================================================
' HELPER UTILITY FUNCTIONS - TYPE-SAFE CONVERSIONS
' ==============================================================================

' ===========NzStr - NULL-SAFE STRING CONVERSION=============
Private Function NzStr(v As Variant) As String
    On Error Resume Next
    
    ' Handle Empty/Null
    If IsEmpty(v) Or IsNull(v) Then
        NzStr = ""
        Exit Function
    End If
    
    ' Handle Errors
    If IsError(v) Then
        NzStr = ""
        Exit Function
    End If
    
    ' Convert to string
    NzStr = CStr(v)
    If Err.Number <> 0 Then NzStr = ""
    
    On Error GoTo 0
End Function



' ===========LoadDefaultPatterns - LOAD PATTERN LOOKUP=============
Private Sub LoadDefaultPatterns()
    On Error GoTo LoadPatErr
    
    ' Find pattern sheet
    Dim patternWS As Worksheet
    On Error Resume Next
    Set patternWS = ThisWorkbook.Sheets("DEFAULT_PATTERN")
    On Error GoTo LoadPatErr
    
    If patternWS Is Nothing Then
        Log "LoadDefaultPatterns: Sheet not found - SKIPPED", "WARNING"
        Exit Sub
    End If
    
    ' Read pattern data
    Dim vPat As Variant
    On Error Resume Next
    vPat = patternWS.UsedRange.Value2
    If Err.Number <> 0 Then
        Log "LoadDefaultPatterns: Failed to read data - " & Err.Description, "ERROR"
        Exit Sub
    End If
    On Error GoTo LoadPatErr
    
    vPat = SafeArray2D(vPat)
    
    ' Build pattern dictionary
    Set g_patternDict = CreateObject("Scripting.Dictionary")
    Dim r As Long, key As String
    Dim rowCount As Long: rowCount = UBound(vPat, 1)
    
    For r = 2 To rowCount ' Skip header
        On Error Resume Next
        
        If Not IsEmpty(vPat(r, 1)) Then
            key = Trim(CStr(vPat(r, 1)))
            If Len(key) > 0 Then
                ' Store pattern value
                g_patternDict(key) = vPat(r, 2)
            End If
        End If
        
        Err.Clear
        On Error GoTo LoadPatErr
    Next r
    
    If g_patternDict.count > 0 Then
        Log "LoadDefaultPatterns: Loaded " & g_patternDict.count & " patterns"
    Else
        Log "LoadDefaultPatterns: No patterns found", "WARNING"
    End If
    
    ' Cleanup
    vPat = Empty
    Exit Sub
    
LoadPatErr:
    Log "LoadDefaultPatterns ERROR: " & Err.Description & " (Err# " & Err.Number & ")", "ERROR"
End Sub

' ==============================================================================
' VALIDATION & TESTING UTILITIES
' ==============================================================================

' ===========ValidateHelperOutput - COMPARE BEFORE/AFTER=============
' USAGE: Call after running helper to validate output matches expected
Public Function ValidateHelperOutput(ws As Worksheet, colName As String, _
                                     expectedSum As Double) As Boolean
    On Error Resume Next
    
    Dim cIdx As Long: cIdx = GetColumnIndex(ws, colName)
    If cIdx = 0 Then
        Log "ValidateHelperOutput: Column '" & colName & "' not found", "ERROR"
        ValidateHelperOutput = False
        Exit Function
    End If
    
    Dim lr As Long: lr = GetLastRow(ws)
    If lr < 2 Then
        ValidateHelperOutput = True
        Exit Function
    End If
    
    ' Calculate sum
    Dim actualSum As Double
    actualSum = Application.WorksheetFunction.Sum(ws.Range(ws.Cells(2, cIdx), ws.Cells(lr, cIdx)))
    
    ' Compare with tolerance
    Dim tolerance As Double: tolerance = 0.01
    Dim diff As Double: diff = Abs(actualSum - expectedSum)
    
    If diff <= tolerance Then
        ValidateHelperOutput = True
        Log "ValidateHelperOutput: '" & colName & "' PASSED (sum=" & Format(actualSum, "0.00") & ")"
    Else
        ValidateHelperOutput = False
        Log "ValidateHelperOutput: '" & colName & "' FAILED (expected=" & expectedSum & ", actual=" & actualSum & ")", "ERROR"
    End If
    
    On Error GoTo 0
End Function

' ===========GetDynamicChunkSize - CALCULATE OPTIMAL CHUNK SIZE=============
Public Function GetDynamicChunkSize(ws As Worksheet) As Long
    On Error Resume Next
    Dim colCount As Long: colCount = GetLastColumn(ws)
    
    ' More aggressive chunking based on column count
    Select Case colCount
        Case Is <= 20: GetDynamicChunkSize = 150000  ' Small datasets
        Case Is <= 40: GetDynamicChunkSize = 100000  ' Medium
        Case Is <= 60: GetDynamicChunkSize = 50000   ' Large
        Case Else: GetDynamicChunkSize = 25000       ' Very large
    End Select
    
    ' Cap based on available memory
    Dim memAvailable As Long: memAvailable = Application.MemoryFree
    If memAvailable < 100000000 Then ' Less than 100MB
        GetDynamicChunkSize = GetDynamicChunkSize \ 2
    End If
    On Error GoTo 0
End Function

' ===========MemoryUsageLog - LOG CURRENT MEMORY USAGE=============
Private Sub MemoryUsageLog(context As String)
    On Error Resume Next
    
    Dim memUsed As Long
    memUsed = Application.MemoryUsed
    
    Log context & " | Memory: " & Format(memUsed / 1024, "#,##0") & " KB"
    
    On Error GoTo 0
End Sub