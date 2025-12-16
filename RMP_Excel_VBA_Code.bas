Option Explicit

' =============================================================================
' RMP EXCEL VBA PROCESSING ENGINE
' =============================================================================
'
' PURPOSE: Process large insurance datasets in batches with modular helper functions
'
' AUTHOR: HKMC Budget Enhancement Team
' VERSION: 2.0
' LAST UPDATED: December 2025
'
' ------------------------------------------------------------------------------
' DEVELOPER GUIDE - AVAILABLE FUNCTIONS FOR HELPER DEVELOPMENT
' ------------------------------------------------------------------------------
'
' COLUMN MANAGEMENT FUNCTIONS:
'   AddColumnIfNotExist(ws, colName)
'     - Adds a new column with specified name if it doesn't exist
'     - Usage: AddColumnIfNotExist myWorksheet, "NEW_COLUMN_NAME"
'
'   GetColumnIndex(ws, colName)
'     - Gets the column index (number) for a given column name
'     - Returns 0 if column doesn't exist
'     - Usage: colIndex = GetColumnIndex(myWorksheet, "EXISTING_COLUMN")
'
' DATA ACCESS FUNCTIONS:
'   NzLong(dataArray, row, col)
'     - Safely converts array element to Long (returns 0 if invalid)
'     - Usage: value = NzLong(myDataArray, rowIndex, columnIndex)
'
'   NzDouble(dataArray, row, col)
'     - Safely converts array element to Double (returns 0# if invalid)
'     - Usage: value = NzDouble(myDataArray, rowIndex, columnIndex)
'
'   NzString(dataArray, row, col)
'     - Safely converts array element to String (returns "" if invalid)
'     - Usage: value = NzString(myDataArray, rowIndex, columnIndex)
'
' ROW/COLUMN INFORMATION FUNCTIONS:
'   GetLastRow(worksheet)
'     - Gets the last row with data in the worksheet
'     - Usage: lastRow = GetLastRow(myWorksheet)
'
'   GetLastColumn(worksheet)
'     - Gets the last column with data in the worksheet
'     - Usage: lastCol = GetLastColumn(myWorksheet)
'
' UTILITY FUNCTIONS:
'   Log(message, optional type)
'     - Logs a message with timestamp (types: INFO, WARNING, ERROR, SUCCESS)
'     - Usage: Log "Processing completed", "SUCCESS"
'
' EXAMPLE HELPER FUNCTION TEMPLATE:
'   Private Sub Helper_99_MyCustomFunction(ws As Worksheet)
'       On Error GoTo ErrorHandler
'       
'       Dim lr As Long: lr = GetLastRow(ws)
'       Dim myCol As Long: myCol = GetColumnIndex(ws, "MY_DATA_COLUMN")
'       
'       If myCol = 0 Then
'           AddColumnIfNotExist ws, "NEW_RESULT_COLUMN"
'           myCol = GetColumnIndex(ws, "NEW_RESULT_COLUMN")
'       End If
'       
'       ' Your custom logic here
'       
'       Exit Sub
'   ErrorHandler:
'       Log "Helper_99_MyCustomFunction ERROR: " & Err.Description, "ERROR"
'   End Sub
'
' ------------------------------------------------------------------------------

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
    bSuccess = False
    Resume SafeExit
End Sub

' ==============================================================================
' CORE UTILITY FUNCTIONS
' ==============================================================================

Private Sub InitializeGlobals()
    Set g_FSO = CreateObject("Scripting.FileSystemObject")
    Set g_HeaderMap = CreateObject("Scripting.Dictionary")
    Set g_LookupData = CreateObject("Scripting.Dictionary")
    Set g_RowIndex = CreateObject("Scripting.Dictionary")
    Set g_colIndexDict = g_HeaderMap
    
    g_ProcessID = Format(Now, "yyyymmdd_hhmmss")
    
    If Not g_FSO.FolderExists(OUT_PATH) Then g_FSO.CreateFolder OUT_PATH
    If Not g_FSO.FolderExists(LOG_PATH) Then g_FSO.CreateFolder LOG_PATH
    
    Dim logFile As String: logFile = LOG_PATH & "Run_" & g_ProcessID & ".txt"
    Set g_LogStream = g_FSO.CreateTextFile(logFile, True)
    
    Log "Process ID: " & g_ProcessID
    Log "Log file: " & logFile
End Sub

Private Sub CleanUpResources()
    On Error Resume Next
    If Not g_LogStream Is Nothing Then g_LogStream.Close
    Set g_FSO = Nothing
    Set g_HeaderMap = Nothing
    Set g_LookupData = Nothing
    Set g_RowIndex = Nothing
    Set g_colIndexDict = Nothing
    Set g_patternDict = Nothing
    Set g_SourceWB = Nothing
    Set g_SourceWS = Nothing
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
    On Error GoTo ErrorHandler
    If Not g_LogStream Is Nothing Then 
        g_LogStream.WriteLine s
    Else
        ' Fallback to immediate window if log stream unavailable
        Debug.Print s
    End If
    Exit Sub
    
ErrorHandler:
    ' If we can't write to log, at least print to immediate window
    Debug.Print "Logging Error: " & Err.Description & " - Message: " & s
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
' DATA LOADING FUNCTIONS
' ==============================================================================

' =======LOAD SOURCE DATA - UNIVERSAL (CSV + EXCEL)=================
Private Function LoadSourceData() As Boolean
    On Error GoTo ErrorHandler
    
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
    
ErrorHandler:
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
    On Error GoTo ErrorHandler
    
    If ws Is Nothing Then
        GetLastRow = 0
        Exit Function
    End If
    
    ' Method 1: Column A (standard approach)
    Dim lr1 As Long: lr1 = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    
    ' Method 2: UsedRange (more reliable for CSV and modified files)
    Dim lr2 As Long: lr2 = 0
    If Not ws.UsedRange Is Nothing Then
        lr2 = ws.UsedRange.Row + ws.UsedRange.Rows.count - 1
    End If
    
    ' Method 3: Direct UsedRange.Rows.count
    Dim lr3 As Long: lr3 = ws.UsedRange.Rows.count
    
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
            GetLastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
            If GetLastRow < 1 Then GetLastRow = ws.UsedRange.Rows.count
        End If
    End If
    
    ' Ensure minimum 1 if data exists
    If GetLastRow < 1 Then
        If ws.UsedRange.Rows.count > 0 Then
            GetLastRow = ws.UsedRange.Rows.count
        Else
            GetLastRow = 1
        End If
    End If
    
    Exit Function
    
ErrorHandler:
    Log "GetLastRow ERROR: " & Err.Description, "ERROR"
    GetLastRow = 1 ' Return safe default
End Function

' ===========GetLastColumn - ROBUST DETECTION=============
Private Function GetLastColumn(ws As Worksheet) As Long
    On Error GoTo ErrorHandler
    
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
    
    Exit Function
    
ErrorHandler:
    Log "GetLastColumn ERROR: " & Err.Description, "ERROR"
    GetLastColumn = 1 ' Return safe default
End Function

' ==============================================================================
' BATCH PROCESSING FUNCTIONS
' ==============================================================================

' ==============================================================================
' PROCESS BATCHES - OPTIMIZED
' ==============================================================================
Private Sub ProcessBatches()
    On Error GoTo ErrorHandler
    
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
    
ErrorHandler:
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
    On Error GoTo ErrorHandler
    
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
        On Error GoTo ErrorHandler
        
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
    
ErrorHandler:
    Log "ExecuteBatchLoop ERROR at Batch " & batchIdx & ": " & Err.Description, "ERROR"
    Err.Raise Err.Number, "ExecuteBatchLoop", Err.Description
End Sub

' ===========ProcessSingleBatch=============
Private Function ProcessSingleBatch(batchIDs As Object, bIdx As Long) As Boolean
    On Error GoTo ErrorHandler
    
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
    
ErrorHandler:
    Log "BATCH ERROR " & bIdx & ": " & Err.Description, "ERROR"
    Resume CleanBatch
End Function

' Create and return new temp workbook/worksheet
Private Sub CreateTemporaryWorkbook(ByRef wbTemp As Workbook, ByRef wsTemp As Worksheet)
    On Error GoTo ErrorHandler
    
    Log "    Creating temp workbook..."
    Set wbTemp = Workbooks.Add(-4167) ' xlWBATemplate.xlWBATWorksheet
    Set wsTemp = wbTemp.Sheets(1)
    Log "    Temp workbook created successfully"
    Exit Sub
    
ErrorHandler:
    Log "    Error creating temp workbook: " & Err.Description
    Err.Clear
    Set wbTemp = Nothing
    Set wsTemp = Nothing
    MsgBox "Error creating temporary workbook: " & Err.Description & vbCrLf & _
           "Please ensure Excel has sufficient resources and permissions.", vbCritical, "Error"
End Sub

' Call FilterDataToSheet for current batch
Private Function FilterBatchData(wsTemp As Worksheet, batchIDs As Object) As Boolean
    On Error GoTo FilterErr
    
    FilterBatchData = False
    
    FilterDataToSheet g_SourceWS, wsTemp, batchIDs
    FilterBatchData = True
    
    Exit Function
    
FilterErr:
    Log "FILTER ERROR: " & Err.Description, "ERROR"
End Function

' Execute all batch helpers
Private Sub ExecuteBatchHelpers(ws As Worksheet)
    RunAllHelpers ws
End Sub

' Save batch output to file
Private Function SaveBatchOutput(wbTemp As Workbook, bIdx As Long) As Boolean
    On Error GoTo ErrorHandler
    
    SaveBatchOutput = False
    
    Dim fName As String
    fName = OUT_PATH & g_FSO.GetBaseName(g_SourceWB.name) & _
            "_Batch_" & Format(bIdx, "000") & "_" & g_ProcessID & ".csv"
    
    Application.DisplayAlerts = False
    wbTemp.SaveAs fName, xlCSV
    Application.DisplayAlerts = True
    
    Log "    Saved: " & g_FSO.GetFileName(fName)
    SaveBatchOutput = True
    
    Exit Function
    
ErrorHandler:
    Log "SAVE ERROR: " & Err.Description, "ERROR"
End Function

' ==============================================================================
' HELPER SYSTEM FUNCTIONS
' ==============================================================================

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

Private Sub FilterDataToSheet(srcWS As Worksheet, destWS As Worksheet, validIDs As Object)
    On Error GoTo FilterErr
    
    If srcWS Is Nothing Then Err.Raise 91, , "Source worksheet is Nothing"
    
    Dim lr As Long: lr = GetLastRow(srcWS)
    If lr < 1 Then
        Log "      No data in source", "WARNING"
        Exit Sub
    End If
    
    Dim lc As Long: lc = GetLastColumn(srcWS)
    
    ' Copy header
    srcWS.Rows(1).Copy destWS.Rows(1)
    
    If lr < 2 Then Exit Sub
    
    ' Read source
    Dim vSrc As Variant
    On Error Resume Next
    vSrc = srcWS.Range("A1", srcWS.Cells(lr, lc)).Value2
    If Err.Number <> 0 Then
        Log "      ERROR reading source: " & Err.Description, "ERROR"
        Exit Sub
    End If
    On Error GoTo FilterErr
    
    vSrc = SafeArray2D(vSrc)
    Dim srcRows As Long: srcRows = UBound(vSrc, 1)
    Dim srcCols As Long: srcCols = UBound(vSrc, 2)
    
    Dim vRes As Variant: ReDim vRes(1 To srcRows, 1 To srcCols)
    Dim outR As Long, r As Long, c As Long
    Dim miNo As Variant, rowCol As Collection, rowNum As Variant
    
    ' Use index if available
    If Not g_RowIndex Is Nothing And g_RowIndex.count > 0 Then
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
        ' Fallback: Full scan
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
    
    If outR > 0 Then
        destWS.Range("A2").Resize(outR, srcCols).Value2 = vRes
    End If
    
    vSrc = Empty
    vRes = Empty
    Set rowCol = Nothing
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
' COLUMN MANAGEMENT FUNCTIONS
' ==============================================================================

' ================= ENHANCED AddColumnIfNotExist - HANDLES ALL SITUATIONS =================
Public Sub AddColumnIfNotExist(ws As Worksheet, colName As String)
    On Error GoTo ErrHandler
    
    Dim ucName As String
    ucName = UCase(Trim(Replace(colName, Chr(160), " "))) ' Normalize
    
    ' Validate input
    If Len(ucName) = 0 Then
        Log "AddColumnIfNotExist: Empty column name provided", "WARNING"
        Exit Sub
    End If
    
    ' Step 1: Refresh header map to avoid stale state
    On Error Resume Next
    RefreshHeaderMap ws
    If Err.Number <> 0 Then
        Log "AddColumnIfNotExist: ERROR in RefreshHeaderMap: " & Err.Description, "WARNING"
        Err.Clear
        ' Continue anyway - try to add column
    End If
    On Error GoTo ErrHandler
    
    ' Step 2: Check if column already exists
    If g_HeaderMap.Exists(ucName) Then
        ' Column exists, verify it's valid
        On Error Resume Next
        Dim existingCol As Long: existingCol = g_HeaderMap(ucName)
        Dim existingHeader As String: existingHeader = ws.Cells(1, existingCol).Value
        If Err.Number = 0 And UCase(Trim(existingHeader)) = ucName Then
            ' Valid existing column
            Exit Sub
        Else
            ' Stale map entry, remove it
            Log "AddColumnIfNotExist: Stale map entry for " & colName & ", refreshing", "DEBUG"
            g_HeaderMap.Remove ucName
            Err.Clear
        End If
        On Error GoTo ErrHandler
    End If
    
    ' Step 3: Determine last column position SAFELY
    Dim lc As Long: lc = 0
    
    ' Method 1: Use HeaderMap (fastest if available)
    On Error Resume Next
    If g_HeaderMap.count > 0 Then
        Dim maxCol As Long: maxCol = 0
        Dim key As Variant
        For Each key In g_HeaderMap.Keys
            Dim colIdx As Long: colIdx = CLng(g_HeaderMap(key))
            If Err.Number = 0 And colIdx > maxCol Then maxCol = colIdx
            Err.Clear
        Next key
        If maxCol > 0 Then lc = maxCol
    End If
    Err.Clear
    On Error GoTo ErrHandler
    
    ' Method 2: Use UsedRange (fallback)
    If lc = 0 Then
        On Error Resume Next
        lc = ws.UsedRange.Columns.count
        If Err.Number <> 0 Or lc = 0 Then
            Err.Clear
            ' Method 3: Scan row 1 (last resort)
            lc = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
            If Err.Number <> 0 Then lc = 1: Err.Clear
        End If
        On Error GoTo ErrHandler
    End If
    
    ' Validate lc is reasonable
    If lc < 1 Then lc = 1
    If lc > 16384 Then lc = 16384 ' Excel column limit
    
    ' Step 4: Find first empty column (in case there are gaps)
    Dim targetCol As Long: targetCol = lc + 1
    On Error Resume Next
    Dim checkAttempts As Long: checkAttempts = 0
    Do While checkAttempts < 100 ' Prevent infinite loop
        Dim headerVal As Variant: headerVal = ws.Cells(1, targetCol).Value
        If Err.Number <> 0 Then
            Err.Clear
            Exit Do
        End If
        
        ' Check if cell is truly empty
        If IsEmpty(headerVal) Or headerVal = "" Or IsNull(headerVal) Then
            Exit Do
        End If
        
        ' Check if this column name matches what we want (case-insensitive)
        If UCase(Trim(CStr(headerVal))) = ucName Then
            ' Column already exists at this position!
            g_HeaderMap(ucName) = targetCol
            Exit Sub
        End If
        
        targetCol = targetCol + 1
        checkAttempts = checkAttempts + 1
    Loop
    Err.Clear
    On Error GoTo ErrHandler
    
    ' Validate target column
    If targetCol > 16384 Then
        Log "AddColumnIfNotExist: Cannot add column " & colName & " - Excel column limit reached", "ERROR"
        Exit Sub
    End If
    
    ' Step 5: Add header with error handling
    On Error Resume Next
    ws.Cells(1, targetCol).Value = colName
    If Err.Number <> 0 Then
        Log "AddColumnIfNotExist: ERROR writing header to column " & targetCol & ": " & Err.Description, "ERROR"
        Err.Clear
        Exit Sub
    End If
    On Error GoTo ErrHandler
    
    ' Step 6: Update header map
    On Error Resume Next
    g_HeaderMap(ucName) = targetCol
    If Err.Number <> 0 Then
        Log "AddColumnIfNotExist: ERROR updating HeaderMap: " & Err.Description, "WARNING"
        Err.Clear
    End If
    On Error GoTo ErrHandler
    
    ' Step 7: Verify addition
    On Error Resume Next
    Dim verifyHeader As String: verifyHeader = ws.Cells(1, targetCol).Value
    If Err.Number = 0 And UCase(Trim(verifyHeader)) = ucName Then
        ' Success - no log needed for normal operation
    Else
        Log "AddColumnIfNotExist: WARNING - Column " & colName & " may not have been added correctly", "WARNING"
    End If
    Err.Clear
    On Error GoTo ErrHandler
    
    Exit Sub

ErrHandler:
    Log "AddColumnIfNotExist: CRITICAL ERROR for column " & colName & ": " & Err.Description & " (Err#" & Err.Number & ")", "ERROR"
    
    ' Attempt recovery: try direct write to next available column
    On Error Resume Next
    Dim recoveryCol As Long: recoveryCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column + 1
    If recoveryCol > 0 And recoveryCol <= 16384 Then
        ws.Cells(1, recoveryCol).Value = colName
        If Err.Number = 0 Then
            g_HeaderMap(ucName) = recoveryCol
            Log "AddColumnIfNotExist: Recovery successful - added " & colName & " at column " & recoveryCol, "WARNING"
        End If
    End If
    Err.Clear
    On Error GoTo 0
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
    vHead = ws.Rows(1).Value2
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

' ==============================================================================
' TYPE-SAFE CONVERSION FUNCTIONS
' ==============================================================================

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

' ========ParseDateFast - HANDLES YYYYMMDD FORMAT===========
Public Function ParseDateFast(ByVal v As Variant, ByRef outDate As Date) As Boolean
    On Error Resume Next
    
    ' Handle Empty/Null/Error
    If IsEmpty(v) Or IsNull(v) Or IsError(v) Then Exit Function
    
    ' Numeric branch
    If IsNumeric(v) Then
        Dim dblVal As Double: dblVal = CDbl(v)
        
        ' YYYYMMDD format
        If dblVal >= 19000101 And dblVal <= 99991231 Then
            Dim y As Long, m As Long, d As Long
            y = Int(dblVal / 10000)
            m = Int((dblVal - y * 10000) / 100)
            d = dblVal - y * 10000 - m * 100
            If y >= 1900 And y <= 9999 And m >= 1 And m <= 12 And d >= 1 And d <= 31 Then
                outDate = DateSerial(y, m, d)
                If Err.Number = 0 Then ParseDateFast = True
            End If
        ' Excel serial
        ElseIf dblVal > 0 And dblVal < 2958466 Then
            outDate = CDate(dblVal)
            If Err.Number = 0 Then ParseDateFast = True
        End If
    
    ' Text branch (CSV often loads dates as text)
    ElseIf VarType(v) = vbString Then
        If IsDate(v) Then
            outDate = CDate(v)
            If Err.Number = 0 Then ParseDateFast = True
        End If
    End If
    
    Err.Clear
    On Error GoTo 0
End Function

' ========ParseTenorFast - UNCHANGED===========
Private Function ParseTenorFast(ByVal v As Variant, ByRef outTenor As Long) As Boolean
    On Error Resume Next
    
    ' Handle Empty/Null/Error
    If IsEmpty(v) Or IsNull(v) Or IsError(v) Then Exit Function
    
    ' Try numeric conversion
    If IsNumeric(v) Then
        outTenor = CLng(v)
        If Err.Number = 0 Then
            ParseTenorFast = True
        End If
    End If
    
    Err.Clear
    On Error GoTo 0
End Function

' ========ADD Column - MAX PERFORMANCE===========
Private Function AddCol(ws As Worksheet, colName As String) As Long
    Dim upperName As String: upperName = UCase(colName)
    
    If g_HeaderMap.Exists(upperName) Then
        AddCol = g_HeaderMap(upperName)
    Else
        Dim lc As Long: lc = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
        If lc = 0 Then lc = 1
        If ws.Cells(1, lc).Value <> "" Then lc = lc + 1
        
        ws.Cells(1, lc).Value = colName
        g_HeaderMap(upperName) = lc
        AddCol = lc
    End If
End Function

' ========MonthDifference - OPTIMIZED===========
Private Function MonthDifference(dateSerial1 As Double, dateSerial2 As Double) As Long
    On Error GoTo CalcErr
    
    Dim d1 As Date, d2 As Date
    d1 = CDate(dateSerial1)
    d2 = CDate(dateSerial2)
    
    MonthDifference = (Year(d2) - Year(d1)) * 12 + (Month(d2) - Month(d1))
    Exit Function
    
CalcErr:
    MonthDifference = -999 ' Error indicator (explicitly handled in caller)
End Function

' =================== OPTIMIZED DATE FUNCTIONS ===================
Private Function ConvertNumericToDate(numericValue As Double) As Date
    ' YYYYMMDD format
    If numericValue >= 19000101 And numericValue <= 99991231 Then
        Dim yearPart As Long: yearPart = Int(numericValue / 10000)
        Dim monthPart As Long: monthPart = Int((numericValue - yearPart * 10000) / 100)
        Dim dayPart As Long: dayPart = numericValue - yearPart * 10000 - monthPart * 100
        ConvertNumericToDate = DateSerial(yearPart, monthPart, dayPart)
    ' Excel serial date
    ElseIf numericValue > 0 And numericValue < 2958466 Then
        ConvertNumericToDate = CDate(numericValue)
    Else
        Err.Raise 13, "ConvertNumericToDate", "Invalid numeric date: " & numericValue
    End If
End Function

Private Function DateDiffMonths(startDate As Date, endDate As Date) As Long
    DateDiffMonths = DateDiff("m", startDate, endDate)
End Function

Private Function DateAddMonths(baseDate As Date, monthsToAdd As Long) As Date
    DateAddMonths = DateAdd("m", monthsToAdd, baseDate)
End Function

' =================== SAFE ARRAY WRAPPER ===================
Private Function SafeArray2D(inputValue As Variant) As Variant
    On Error Resume Next
    
    If Not IsArray(inputValue) Then
        Dim singleValueArray(1 To 1, 1 To 1) As Variant
        singleValueArray(1, 1) = inputValue
        SafeArray2D = singleValueArray
        Exit Function
    End If
    
    Dim testRows As Long: testRows = UBound(inputValue, 1)
    Dim testCols As Long: testCols = UBound(inputValue, 2)
    
    If Err.Number = 0 Then
        SafeArray2D = inputValue
    Else
        Err.Clear
        Dim arrayBound As Long: arrayBound = UBound(inputValue)
        Dim converted2D(1 To 1, 1 To 1) As Variant
        If arrayBound >= LBound(inputValue) Then
            converted2D(1, 1) = inputValue(LBound(inputValue))
        End If
        SafeArray2D = converted2D
    End If
    
    On Error GoTo 0
End Function

' =================== FAST CONVERSION (MATCHES NzLong) ===================
Private Function FastCLng(v As Variant) As Long
    Select Case VarType(v)
        Case 2 To 6, 14, 17, 20  ' Numeric types
            On Error Resume Next
            FastCLng = CLng(v)
            If Err.Number <> 0 Then FastCLng = 0: Err.Clear
        Case 8  ' String
            On Error Resume Next
            FastCLng = CLng(v)
            If Err.Number <> 0 Then FastCLng = 0: Err.Clear
        Case Else
            FastCLng = 0
    End Select
End Function

' =================== NZD - OPTIMIZED ===================
Private Function NzD(v As Variant) As Double
    '  MATCHES OLD: Error-safe conversion with overflow protection
    On Error Resume Next
    If IsNumeric(v) Then NzD = CDbl(v)
    If Err.Number <> 0 Then NzD = 0#: Err.Clear
End Function

' =================== PDATEFAST - OPTIMIZED ===================
Private Function PDateFast(v As Variant, ByRef outDate As Date) As Boolean
    '  MATCHES OLD: Robust date parsing with error handling
    On Error Resume Next
    
    If IsEmpty(v) Or IsNull(v) Or IsError(v) Then Exit Function
    
    If IsNumeric(v) Then
        Dim dv As Double: dv = CDbl(v)
        
        ' YYYYMMDD format
        If dv >= 19000101# And dv <= 99991231# Then
            Dim y As Long, m As Long, d As Long
            y = Int(dv / 10000)
            m = Int((dv - y * 10000) / 100)
            d = dv - y * 10000 - m * 100
            If y >= 1900 And y <= 9999 And m >= 1 And m <= 12 And d >= 1 And d <= 31 Then
                outDate = DateSerial(y, m, d)
                PDateFast = (Err.Number = 0)
            End If
        ' Excel serial (>0 to reject serial 0)
        ElseIf dv > 0# And dv < 2958466# Then
            outDate = CDate(dv)
            PDateFast = (Err.Number = 0)
        End If
    
    ' String branch
    ElseIf VarType(v) = 8 Then
        If IsDate(v) Then
            outDate = CDate(v)
            PDateFast = (Err.Number = 0)
        End If
    End If
    
    Err.Clear
End Function

' =================== INLINE NZDOUBLE (OPTIMIZED) ===================
Private Function InlineNzDouble(v As Variant) As Double
    On Error Resume Next
    Select Case VarType(v)
        Case 2 To 6, 14, 17, 20: InlineNzDouble = CDbl(v)
        Case 8: InlineNzDouble = CDbl(v)
        Case Else: InlineNzDouble = 0#
    End Select
    If Err.Number <> 0 Then InlineNzDouble = 0#: Err.Clear
End Function

' =================== ProcessChunked ===================
Private Sub ProcessSinglePass(ws As Worksheet, lr As Long, _
                              cOpt As Long, cAmt As Long, cRate As Long, cRIPct As Long, _
                              cYr As Long, cMon As Long, cDur As Long, cRI As Long, _
                              cComm As Long, cCommBOP As Long, cRIComm As Long, cRICommBOP As Long)
    
    Dim vOpt As Variant: vOpt = ws.Range(ws.Cells(2, cOpt), ws.Cells(lr, cOpt)).Value2
    Dim vAmt As Variant: vAmt = ws.Range(ws.Cells(2, cAmt), ws.Cells(lr, cAmt)).Value2
    Dim vRate As Variant: vRate = ws.Range(ws.Cells(2, cRate), ws.Cells(lr, cRate)).Value2
    Dim vRIPct As Variant: vRIPct = ws.Range(ws.Cells(2, cRIPct), ws.Cells(lr, cRIPct)).Value2
    Dim vYr As Variant: vYr = ws.Range(ws.Cells(2, cYr), ws.Cells(lr, cYr)).Value2
    Dim vMon As Variant: vMon = ws.Range(ws.Cells(2, cMon), ws.Cells(lr, cMon)).Value2
    Dim vDur As Variant: vDur = ws.Range(ws.Cells(2, cDur), ws.Cells(lr, cDur)).Value2
    Dim vRI As Variant: vRI = ws.Range(ws.Cells(2, cRI), ws.Cells(lr, cRI)).Value2
    
    vOpt = SafeArray2D(vOpt): vAmt = SafeArray2D(vAmt): vRate = SafeArray2D(vRate)
    vRIPct = SafeArray2D(vRIPct): vYr = SafeArray2D(vYr): vMon = SafeArray2D(vMon)
    vDur = SafeArray2D(vDur): vRI = SafeArray2D(vRI)
    
    Dim rows As Long: rows = UBound(vOpt, 1)
    Dim vOut() As Double: ReDim vOut(1 To rows, 1 To 4)
    
    Dim r As Long, v As Variant, optType As String
    Dim amt As Double, rate As Double, riPct As Double, calcComm As Double
    Dim yr As Long, mon As Long, dur As Long, isRI As Boolean
    Dim errCount As Long: errCount = 0
    
    For r = 1 To rows
        v = vOpt(r, 1)
        If VarType(v) = vbString Then
            optType = UCase$(Trim$(CStr(v)))
            
            If optType = "SI" Then
                amt = 0: rate = 0: riPct = 0: yr = 0: mon = 0: dur = 0: isRI = False
                
                v = vAmt(r, 1)
                If IsNumeric(v) And Not IsEmpty(v) Then
                    amt = CDbl(v)
                Else
                    errCount = errCount + 1
                End If
                
                v = vRate(r, 1)
                If IsNumeric(v) And Not IsEmpty(v) Then
                    rate = CDbl(v)
                Else
                    errCount = errCount + 1
                End If
                
                v = vRIPct(r, 1)
                If IsNumeric(v) And Not IsEmpty(v) Then
                    riPct = CDbl(v) * 0.01
                Else
                    errCount = errCount + 1
                End If
                
                v = vYr(r, 1)
                If IsNumeric(v) And Not IsEmpty(v) Then yr = CLng(v)
                
                v = vMon(r, 1)
                If IsNumeric(v) And Not IsEmpty(v) Then mon = CLng(v)
                
                v = vDur(r, 1)
                If IsNumeric(v) And Not IsEmpty(v) Then dur = CLng(v)
                
                v = vRI(r, 1)
                If VarType(v) = vbString Then
                    isRI = (UCase$(Trim$(CStr(v))) = "Y")
                ElseIf IsNumeric(v) Then
                    isRI = (CLng(v) = 1)
                End If
                
                calcComm = amt * rate
                
                If yr = 1 And mon = 1 Then
                    vOut(r, 1) = calcComm
                    If isRI Then vOut(r, 3) = calcComm * riPct
                End If
                
                If dur = 1 Then
                    vOut(r, 2) = calcComm
                    If isRI Then vOut(r, 4) = calcComm * riPct
                End If
            End If
        End If
    Next r
    
    ws.Cells(2, cComm).Resize(rows, 4).Value2 = vOut
    
    Erase vOpt, vAmt, vRate, vRIPct, vYr, vMon, vDur, vRI, vOut
    
    If errCount > 0 Then Log "H28: " & errCount & " conversion errors", "WARNING"
End Sub

' =================== ProcessChunked ===================
Private Sub ProcessChunked(ws As Worksheet, lr As Long, _
                           cOpt As Long, cAmt As Long, cRate As Long, cRIPct As Long, _
                           cYr As Long, cMon As Long, cDur As Long, cRI As Long, _
                           cComm As Long, cCommBOP As Long, cRIComm As Long, cRICommBOP As Long, _
                           is64Bit As Boolean)
    
    Dim targetMB As Long: targetMB = IIf(is64Bit, 75, 30)
    Dim chunkRows As Long: chunkRows = (targetMB * 131072) \ 8
    If chunkRows > 150000 Then chunkRows = 150000
    If chunkRows < 5000 Then chunkRows = 5000
    
    Dim chunkStart As Long: chunkStart = 2
    Dim chunkEnd As Long, rows As Long, r As Long
    Dim vOpt As Variant, vAmt As Variant, vRate As Variant, vRIPct As Variant
    Dim vYr As Variant, vMon As Variant, vDur As Variant, vRI As Variant
    Dim vOut() As Double, v As Variant, optType As String
    Dim amt As Double, rate As Double, riPct As Double, calcComm As Double
    Dim yr As Long, mon As Long, dur As Long, isRI As Boolean
    Dim totalErrors As Long: totalErrors = 0
    
    Do While chunkStart <= lr
        chunkEnd = Application.Min(chunkStart + chunkRows - 1, lr)
        
        vOpt = SafeArray2D(ws.Range(ws.Cells(chunkStart, cOpt), ws.Cells(chunkEnd, cOpt)).Value2)
        vAmt = SafeArray2D(ws.Range(ws.Cells(chunkStart, cAmt), ws.Cells(chunkEnd, cAmt)).Value2)
        vRate = SafeArray2D(ws.Range(ws.Cells(chunkStart, cRate), ws.Cells(chunkEnd, cRate)).Value2)
        vRIPct = SafeArray2D(ws.Range(ws.Cells(chunkStart, cRIPct), ws.Cells(chunkEnd, cRIPct)).Value2)
        vYr = SafeArray2D(ws.Range(ws.Cells(chunkStart, cYr), ws.Cells(chunkEnd, cYr)).Value2)
        vMon = SafeArray2D(ws.Range(ws.Cells(chunkStart, cMon), ws.Cells(chunkEnd, cMon)).Value2)
        vDur = SafeArray2D(ws.Range(ws.Cells(chunkStart, cDur), ws.Cells(chunkEnd, cDur)).Value2)
        vRI = SafeArray2D(ws.Range(ws.Cells(chunkStart, cRI), ws.Cells(chunkEnd, cRI)).Value2)
        
        rows = UBound(vOpt, 1)
        ReDim vOut(1 To rows, 1 To 4)
        
        For r = 1 To rows
            v = vOpt(r, 1)
            If VarType(v) = vbString Then
                optType = UCase$(Trim$(CStr(v)))
                
                If optType = "SI" Then
                    amt = 0: rate = 0: riPct = 0: yr = 0: mon = 0: dur = 0: isRI = False
                    
                    v = vAmt(r, 1)
                    If IsNumeric(v) And Not IsEmpty(v) Then
                        amt = CDbl(v)
                    Else
                        totalErrors = totalErrors + 1
                    End If
                    
                    v = vRate(r, 1)
                    If IsNumeric(v) And Not IsEmpty(v) Then
                        rate = CDbl(v)
                    Else
                        totalErrors = totalErrors + 1
                    End If
                    
                    v = vRIPct(r, 1)
                    If IsNumeric(v) And Not IsEmpty(v) Then
                        riPct = CDbl(v) * 0.01
                    Else
                        totalErrors = totalErrors + 1
                    End If
                    
                    v = vYr(r, 1)
                    If IsNumeric(v) And Not IsEmpty(v) Then yr = CLng(v)
                    
                    v = vMon(r, 1)
                    If IsNumeric(v) And Not IsEmpty(v) Then mon = CLng(v)
                    
                    v = vDur(r, 1)
                    If IsNumeric(v) And Not IsEmpty(v) Then dur = CLng(v)
                    
                    v = vRI(r, 1)
                    If VarType(v) = vbString Then
                        isRI = (UCase$(Trim$(CStr(v))) = "Y")
                    ElseIf IsNumeric(v) Then
                        isRI = (CLng(v) = 1)
                    End If
                    
                    calcComm = amt * rate
                    
                    If yr = 1 And mon = 1 Then
                        vOut(r, 1) = calcComm
                        If isRI Then vOut(r, 3) = calcComm * riPct
                    End If
                    
                    If dur = 1 Then
                        vOut(r, 2) = calcComm
                        If isRI Then vOut(r, 4) = calcComm * riPct
                    End If
                End If
            End If
        Next r
        
        ws.Cells(chunkStart, cComm).Resize(rows, 4).Value2 = vOut
        
        Erase vOpt, vAmt, vRate, vRIPct, vYr, vMon, vDur, vRI, vOut
        
        chunkStart = chunkEnd + 1
    Loop
    
    If totalErrors > 0 Then Log "H28: " & totalErrors & " conversion errors", "WARNING"
End Sub

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
' END OF MAIN MODULE
' ==============================================================================