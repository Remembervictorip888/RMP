Option Explicit

' =============================================================================
' RMP EXCEL VBA PROCESSING ENGINE - ENHANCED VERSION
' =============================================================================
'
' PURPOSE: Process large insurance datasets in batches with modular helper functions
' AUTHOR: HKMC Budget Enhancement Team
' VERSION: 2.3
' LAST UPDATED: December 2025
'
' ENHANCEMENTS MADE:
' 1. Enhanced performance for large-scale data processing
' 2. Eliminated silent failures with explicit error handling
' 3. Implemented fast exit strategies with comprehensive logging
' 4. Improved memory management for large datasets
' 5. Enhanced error resilience with detailed diagnostics
'
' ------------------------------------------------------------------------------
' DEVELOPER GUIDE - AVAILABLE FUNCTIONS FOR HELPER DEVELOPMENT
' ------------------------------------------------------------------------------
'
' COLUMN MANAGEMENT FUNCTIONS:
'   AddColumn(ws, colName)
'     - Adds a new column with specified name if it doesn't exist
'     - Usage: AddColumn myWorksheet, "NEW_COLUMN_NAME"
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
'   SafeArray2D(inputValue)
'     - Wraps variant values in a 2D array for safe processing
'     - Usage: dataArray = SafeArray2D(worksheet.Range(...).Value2)
'
' DATE FUNCTIONS:
'   ParseDate(value, ByRef outDate)
'     - Parses various date formats into Date type
'     - Usage: If ParseDate(cellValue, myDate) Then ...
'
' EXAMPLE HELPER FUNCTION TEMPLATE:
'   Private Sub Helper_99_MyCustomFunction(ws As Worksheet)
'       ' Declare all variables upfront (Clean code)
'       Dim startTime As Double: startTime = Timer
'       Dim lr As Long, lc As Long
'       Dim cRequiredCol1 As Long, cRequiredCol2 As Long
'       Dim cOutputCol As Long
'       Dim vDataArray As Variant
'       Dim i As Long
'       
'       Log "Helper_99_MyCustomFunction START"
'       
'       On Error GoTo ErrorHandler
'       
'       ' Defensive programming - Validate columns and data types upfront
'       cRequiredCol1 = GetColumnIndex(ws, "REQUIRED_COL1")
'       If cRequiredCol1 = 0 Then
'           Log "Helper_99_MyCustomFunction ERROR: Required column 'REQUIRED_COL1' not found", "ERROR"
'           Exit Sub ' Fail fast
'       End If
'       
'       cRequiredCol2 = GetColumnIndex(ws, "REQUIRED_COL2")
'       If cRequiredCol2 = 0 Then
'           Log "Helper_99_MyCustomFunction ERROR: Required column 'REQUIRED_COL2' not found", "ERROR"
'           Exit Sub ' Fail fast
'       End If
'       
'       ' Get data dimensions
'       lr = GetLastRow(ws)
'       lc = GetLastColumn(ws)
'       If lr < 2 Then
'           Log "Helper_99_MyCustomFunction WARNING: No data rows to process", "WARNING"
'           Exit Sub ' Fail fast
'       End If
'       
'       ' Speed/Scalability - Dynamic chunking based on column count
'       Dim chunkSize As Long
'       chunkSize = CalculateOptimalChunkSize(lc, lr)
'       
'       ' Speed/Scalability - In-memory arrays with SafeArray2D
'       vDataArray = SafeArray2D(ws.Range(ws.Cells(1, 1), ws.Cells(lr, lc)).Value2)
'       
'       ' Modular - Reuse shared utilities (no duplicates)
'       AddColumn ws, "OUTPUT_COLUMN"
'       cOutputCol = GetColumnIndex(ws, "OUTPUT_COLUMN")
'       
'       ' Process data in chunks for scalability
'       Dim chunkStart As Long: chunkStart = 2
'       Dim chunkEnd As Long
'       
'       Do While chunkStart <= lr
'           chunkEnd = Application.Min(chunkStart + chunkSize - 1, lr)
'           
'           ' Your processing logic here using vDataArray
'           For i = chunkStart To chunkEnd
'               ' Process each row
'               ' Use NzLong, NzDouble, NzString for safe data access
'           Next i
'           
'           chunkStart = chunkEnd + 1
'       Loop
'       
'       ' Validate output integrity (Defensive)
'       ' Example: Check if all mappings were successful
'       ' If validation fails, use Exit Sub to fail early
'       
'       ' Record successful completion with duration
'       Dim duration As Double: duration = Timer - startTime
'       Log "Helper_99_MyCustomFunction COMPLETED | Duration: " & Format(duration, "0.000") & "s", "SUCCESS"
'       Exit Sub
'       
'   ErrorHandler:
'       ' No silent errors - Detailed log on failure
'       Dim errorDuration As Double: errorDuration = Timer - startTime
'       Log "Helper_99_MyCustomFunction FAILED | Duration: " & Format(errorDuration, "0.000") & "s | Error: " & Err.Description, "ERROR"
'   End Sub
'
' BEST PRACTICES FOR HELPER DEVELOPMENT:
'   1. FAIL FAST: Exit immediately with detailed logs on any failure
'   2. DEFENSIVE PROGRAMMING: Validate columns, data types, and integrity upfront
'   3. MODULARITY: Reuse shared utilities (no code duplication)
'   4. SPEED/SCALABILITY: 
'      - Use in-memory arrays with SafeArray2D()
'      - Pre-cache column indices
'      - Read/write once
'      - Apply dynamic chunking based on system resources
'      - Disable ScreenUpdating/Calculation/DisplayAlerts (already handled by engine)
'   5. CLEAN CODE:
'      - KISS/DRY principles
'      - Meaningful variable names (e.g., cColName, vDataArray)
'      - Declare all variables upfront
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
    On Error GoTo ErrorHandler
    
    Set g_FSO = CreateObject("Scripting.FileSystemObject")
    Set g_HeaderMap = CreateObject("Scripting.Dictionary")
    Set g_LookupData = CreateObject("Scripting.Dictionary")
    Set g_RowIndex = CreateObject("Scripting.Dictionary")
    Set g_colIndexDict = g_HeaderMap
    
    ' Pre-size dictionaries for better performance
    g_HeaderMap.CompareMode = vbTextCompare
    g_RowIndex.CompareMode = vbTextCompare
    g_LookupData.CompareMode = vbTextCompare
    
    g_ProcessID = Format(Now, "yyyymmdd_hhmmss")
    
    If Not g_FSO.FolderExists(OUT_PATH) Then 
        On Error Resume Next
        g_FSO.CreateFolder OUT_PATH
        If Err.Number <> 0 Then
            Log "Failed to create Output folder: " & Err.Description, "ERROR"
            Err.Clear
        End If
        On Error GoTo ErrorHandler
    End If
    
    If Not g_FSO.FolderExists(LOG_PATH) Then 
        On Error Resume Next
        g_FSO.CreateFolder LOG_PATH
        If Err.Number <> 0 Then
            Log "Failed to create Log folder: " & Err.Description, "ERROR"
            Err.Clear
        End If
        On Error GoTo ErrorHandler
    End If
    
    Dim logFile As String: logFile = LOG_PATH & "Run_" & g_ProcessID & ".txt"
    On Error Resume Next
    Set g_LogStream = g_FSO.CreateTextFile(logFile, True)
    If Err.Number <> 0 Then
        Log "Failed to create log file: " & Err.Description, "ERROR"
        Err.Clear
    End If
    On Error GoTo ErrorHandler
    
    Log "Process ID: " & g_ProcessID
    Log "Log file: " & logFile
    Exit Sub
    
ErrorHandler:
    Log "InitializeGlobals ERROR: " & Err.Description & " (Err# " & Err.Number & ")", "ERROR"
End Sub

Private Sub CleanUpResources()
    On Error Resume Next
    If Not g_LogStream Is Nothing Then 
        g_LogStream.Close
        Set g_LogStream = Nothing
    End If
    Set g_FSO = Nothing
    Set g_HeaderMap = Nothing
    Set g_LookupData = Nothing
    Set g_RowIndex = Nothing
    Set g_colIndexDict = Nothing
    Set g_patternDict = Nothing
    If Not g_SourceWB Is Nothing Then
        g_SourceWB.Close False
        Set g_SourceWB = Nothing
    End If
    Set g_SourceWS = Nothing
    On Error GoTo 0
End Sub

Private Sub ToggleOptimization(bOn As Boolean)
    On Error Resume Next
    With Application
        .ScreenUpdating = Not bOn
        .Calculation = IIf(bOn, xlCalculationManual, xlCalculationAutomatic)
        .EnableEvents = Not bOn
        .DisplayAlerts = Not bOn
        .StatusBar = Not bOn
    End With
    On Error GoTo 0
End Sub

Private Sub Log(msg As String, Optional sType As String = "INFO")
    Dim s As String: s = Format(Now, "hh:mm:ss") & " [" & sType & "] " & msg
    
    ' Print to immediate window if DEBUG_PRINT is True
    If DEBUG_PRINT Then Debug.Print s
    
    ' Always attempt to write to log file
    On Error Resume Next
    If Not g_LogStream Is Nothing Then
        g_LogStream.WriteLine s
        ' Force write to disk for critical errors
        If sType = "ERROR" Or sType = "FATAL" Then
            g_LogStream.Flush
        End If
    End If
    On Error GoTo 0
End Sub

' Calculate optimal chunk size based on system resources and data characteristics
Private Function CalculateOptimalChunkSize(colCount As Long, rowCount As Long) As Long
    On Error GoTo ErrorHandler
    
    ' Base calculation on available memory and data density
    Dim baseSize As Long
    
    ' Adjust chunk size based on column count (fewer columns = larger chunks)
    If colCount <= 30 Then
        baseSize = 100000 ' ≤30 cols = 100k rows
    ElseIf colCount <= 50 Then
        baseSize = 50000  ' 30-50 cols = 50k rows
    Else
        baseSize = 25000  ' >50 cols = 25k rows
    End If
    
    ' Further adjust based on system architecture (64-bit can handle more)
    #If Win64 Then
        baseSize = baseSize * 1.5
    #End If
    
    ' Ensure we don't exceed the total row count
    CalculateOptimalChunkSize = Application.Min(baseSize, rowCount)
    
    ' Ensure minimum chunk size
    If CalculateOptimalChunkSize < 1000 Then CalculateOptimalChunkSize = 1000
    
    Exit Function
    
ErrorHandler:
    ' Default fallback values
    CalculateOptimalChunkSize = 10000
    Log "CalculateOptimalChunkSize ERROR: " & Err.Description & " - Using default chunk size", "WARNING"
End Function

Private Function BrowseForFile() As String
    On Error GoTo ErrorHandler
    
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "Select Input File"
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xlsx;*.xls;*.xlsm"
        .Filters.Add "CSV Files", "*.csv"
        .Filters.Add "All Files", "*.*"
        If .Show = -1 Then 
            BrowseForFile = .SelectedItems(1)
        Else
            BrowseForFile = ""
        End If
    End With
    
    Exit Function
    
ErrorHandler:
    Log "BrowseForFile ERROR: " & Err.Description, "ERROR"
    BrowseForFile = ""
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
    Log "File type detected: " & UCase$(fileExt)
    
    ' Open source workbook
    If Not OpenSourceWorkbook(fPath, fileExt) Then
        Log "Failed to open workbook", "ERROR"
        Exit Function
    End If
    
    ' Validate workbook opened
    If g_SourceWB Is Nothing Then
        Log "Workbook object is Nothing", "ERROR"
        Exit Function
    End If
    
    ' Validate source worksheet
    If Not ValidateSourceWorksheet() Then
        Log "Source worksheet validation failed", "ERROR"
        Exit Function
    End If
    
    ' Get dimensions
    Dim lr As Long: lr = GetLastRow(g_SourceWS)
    Dim lc As Long: lc = GetLastColumn(g_SourceWS)
    
    Log "Loaded: " & lr & " rows × " & lc & " columns"
    
    ' Validate minimum structure
    If lr < 1 Then
        Log "No rows found", "ERROR"
        Exit Function
    End If
    
    If lc < 1 Then
        Log "No columns found", "ERROR"
        Exit Function
    End If
    
    ' Build header map
    BuildSourceHeaderMap g_SourceWS
    
    LoadSourceData = True
    Exit Function
    
ErrorHandler:
    Log "LoadSourceData ERROR: " & Err.Description & " (Err# " & Err.Number & ")", "ERROR"
    LoadSourceData = False
    
Finally:
    On Error Resume Next
    If Not g_SourceWB Is Nothing And LoadSourceData = False Then 
        g_SourceWB.Close False
        Set g_SourceWS = Nothing
        Set g_SourceWB = Nothing
    End If
    On Error GoTo 0
End Function

' Prompt for source file using file dialog
Private Function PromptForSourceFile() As String
    On Error GoTo ErrorHandler
    PromptForSourceFile = BrowseForFile()
    Exit Function
    
ErrorHandler:
    Log "PromptForSourceFile ERROR: " & Err.Description, "ERROR"
    PromptForSourceFile = ""
End Function

' Validate that the specified file exists
Private Function ValidateFileExists(fPath As String) As Boolean
    On Error GoTo ErrorHandler
    
    If Not g_FSO Is Nothing Then
        ValidateFileExists = g_FSO.FileExists(fPath)
    Else
        ValidateFileExists = False
    End If
    
    Exit Function
    
ErrorHandler:
    Log "ValidateFileExists ERROR: " & Err.Description, "ERROR"
    ValidateFileExists = False
End Function

' Detect and normalize the source file type
Private Function DetectSourceFileType(fPath As String) As String
    On Error GoTo ErrorHandler
    
    If Not g_FSO Is Nothing Then
        DetectSourceFileType = LCase$(g_FSO.GetExtensionName(fPath))
    Else
        DetectSourceFileType = ""
    End If
    
    Exit Function
    
ErrorHandler:
    Log "DetectSourceFileType ERROR: " & Err.Description, "ERROR"
    DetectSourceFileType = ""
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
            OpenSourceWorkbook = True
        Case "xlsx", "xlsm", "xls"
            ' Excel: Standard open
            Set g_SourceWB = Workbooks.Open(FileName:=fPath, _
                                             ReadOnly:=True, _
                                             UpdateLinks:=0)
            OpenSourceWorkbook = True
        Case Else
            Log "Unsupported file type: " & fileExt, "ERROR"
            Exit Function
    End Select
    
    Exit Function
    
OpenErr:
    Log "Failed to open " & UCase$(fileExt) & ": " & Err.Description, "ERROR"
    Set g_SourceWB = Nothing
End Function

' Validate source worksheet and assign first worksheet to g_SourceWS
Private Function ValidateSourceWorksheet() As Boolean
    On Error GoTo ValidateErr
    
    ValidateSourceWorksheet = False
    
    ' Get first worksheet
    If g_SourceWB.Sheets.Count = 0 Then
        Log "Workbook has no sheets", "ERROR"
        GoTo CleanUp
    End If
    
    ' Ensure we're getting a Worksheet object, not a Chart or other sheet type
    Dim i As Integer
    For i = 1 To g_SourceWB.Sheets.Count
        If TypeName(g_SourceWB.Sheets(i)) = "Worksheet" Then
            Set g_SourceWS = g_SourceWB.Sheets(i)
            Exit For
        End If
    Next i
    
    If g_SourceWS Is Nothing Then
        Log "Failed to access worksheet (no Worksheets found in workbook)", "ERROR"
        GoTo CleanUp
    End If
    
    ValidateSourceWorksheet = True

CleanUp:
    If Not ValidateSourceWorksheet Then
        If Not g_SourceWB Is Nothing Then
            g_SourceWB.Close False
            Set g_SourceWB = Nothing
        End If
        Set g_SourceWS = Nothing
    End If
    Exit Function
    
ValidateErr:
    Log "ValidateSourceWorksheet ERROR: " & Err.Description, "ERROR"
    Resume CleanUp
End Function

' Build source header map from source worksheet
Private Sub BuildSourceHeaderMap(ws As Worksheet)
    On Error GoTo ErrorHandler
    RefreshHeaderMap ws
    Exit Sub
    
ErrorHandler:
    Log "BuildSourceHeaderMap ERROR: " & Err.Description, "ERROR"
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
    
    ' Cache UsedRange to avoid repeated property access
    Dim usedRng As Range
    Set usedRng = Nothing
    On Error Resume Next
    Set usedRng = ws.UsedRange
    On Error GoTo ErrorHandler
    
    ' Method 1: Column A (standard approach)
    Dim lr1 As Long: lr1 = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' Method 2: UsedRange (more reliable for CSV and modified files)
    Dim lr2 As Long: lr2 = 0
    If Not usedRng Is Nothing Then
        lr2 = usedRng.Row + usedRng.Rows.Count - 1
    End If
    
    ' Method 3: Direct UsedRange.Rows.Count
    Dim lr3 As Long: lr3 = 0
    If Not usedRng Is Nothing Then
        lr3 = usedRng.Rows.Count
    End If
    
    ' Method 4: Special handling for CSV - check if we have data beyond header
    Dim lr4 As Long: lr4 = 0
    If lr1 <= 1 And lr3 > 1 Then
        ' CSV edge case: Method 1 fails but UsedRange shows data
        lr4 = lr3
    End If
    
    ' Use maximum of all methods
    GetLastRow = Application.Max(lr1, lr2, lr3, lr4)
    
    ' Final validation
    If GetLastRow < 1 Then
        ' Last resort: Count non-empty cells in column A starting from row 2
        Dim cell As Range
        Set cell = ws.Cells(2, 1)
        If Not IsEmpty(cell) And Not IsError(cell) Then
            GetLastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        End If
    End If
    
    ' Ensure reasonable result
    If GetLastRow < 1 Then
        If lr3 > 0 Then
            GetLastRow = lr3
        Else
            GetLastRow = 1
        End If
    End If
    
    Exit Function
    
ErrorHandler:
    Log "GetLastRow ERROR: " & Err.Description & " - returning default value", "ERROR"
    GetLastRow = 1 ' Return safe default
End Function

' ===========GetLastColumn - ROBUST DETECTION=============
Private Function GetLastColumn(ws As Worksheet) As Long
    On Error GoTo ErrorHandler
    
    If ws Is Nothing Then
        GetLastColumn = 0
        Exit Function
    End If
    
    ' Find the last column with data by checking multiple rows
    Dim lastCol As Long
    Dim i As Long
    Dim candidateCol As Long
    
    ' Check first row (header) and a few data rows for last column
    For i = 1 To Application.Min(10, ws.Rows.Count) ' Check first 10 rows or total rows if less
        candidateCol = ws.Cells(i, ws.Columns.Count).End(xlToLeft).Column
        If candidateCol > lastCol Then lastCol = candidateCol
    Next i
    
    ' Also check from bottom up to catch any data in far rows
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow > 10 Then ' If we have more than 10 rows, check a few more
        For i = lastRow To Application.Max(11, lastRow - 5) Step -1
            candidateCol = ws.Cells(i, ws.Columns.Count).End(xlToLeft).Column
            If candidateCol > lastCol Then lastCol = candidateCol
        Next i
    End If
    
    GetLastColumn = lastCol
    
    ' Ensure minimum 1
    If GetLastColumn < 1 Then GetLastColumn = 1
    
    Exit Function
    
ErrorHandler:
    Log "GetLastColumn ERROR: " & Err.Description & " - returning default value", "ERROR"
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
    Dim colUnique As Collection
    Set colUnique = GetDistinctMI_NOs()
    If colUnique Is Nothing Or colUnique.Count = 0 Then
        Log "No distinct MI_NOs found", "WARNING"
        Exit Sub
    End If
    
    Log "Found " & colUnique.Count & " distinct MI_NOs"
    
    ' Build row index for fast filtering
    BuildMI_NORowIndex
    
    ' Calculate optimal batch size
    Dim batchSize As Long
    batchSize = CalculateOptimalBatchSize(colUnique.Count)
    If batchSize <= 0 Then batchSize = 1 ' Ensure minimum batch size
    
    Dim totalBatches As Long
    totalBatches = CalculateTotalBatches(colUnique.Count, batchSize)
    Log "Batch size: " & batchSize & " | Total batches: " & totalBatches
    
    ' Execute batch loop
    ExecuteBatchLoop colUnique, batchSize, totalBatches
    
    Log "All " & totalBatches & " batches processed"
    
    Set colUnique = Nothing
    Exit Sub
    
ErrorHandler:
    If Not colUnique Is Nothing Then Set colUnique = Nothing
    Log "ProcessBatches ERROR: " & Err.Description & " (Err# " & Err.Number & ")", "ERROR"
    Err.Raise Err.Number, "ProcessBatches", Err.Description
End Sub

' Get collection of unique MI_NOs
Private Function GetDistinctMI_NOs() As Collection
    On Error GoTo ErrorHandler
    Set GetDistinctMI_NOs = GetDistinctIDs(g_SourceWS)
    Exit Function
    
ErrorHandler:
    Log "Error in GetDistinctMI_NOs: " & Err.Description & " (Err# " & Err.Number & ")", "ERROR"
    Set GetDistinctMI_NOs = New Collection
End Function

' Build g_RowIndex for fast filtering
Private Sub BuildMI_NORowIndex()
    On Error GoTo ErrorHandler
    Log "Building row index..."
    Dim tIndex As Double: tIndex = Timer
    Application.ScreenUpdating = False
    BuildRowIndex g_SourceWS
    Application.ScreenUpdating = True
    Log "Row index built in " & Format(Timer - tIndex, "0.00") & "s"
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Log "Error in BuildMI_NORowIndex: " & Err.Description & " (Err# " & Err.Number & ")", "ERROR"
End Sub

' Calculate optimal batch size based on total distinct IDs
Private Function CalculateOptimalBatchSize(totalIDs As Long) As Long
    On Error GoTo ErrorHandler
    If totalIDs <= 0 Then
        CalculateOptimalBatchSize = 1
    Else
        CalculateOptimalBatchSize = CalculateBatchSize(totalIDs)
    End If
    Exit Function
    
ErrorHandler:
    Log "Error in CalculateOptimalBatchSize: " & Err.Description & " - using default batch size", "WARNING"
    CalculateOptimalBatchSize = 1000 ' Default batch size
End Function

' Calculate total number of batches
Private Function CalculateTotalBatches(totalIDs As Long, batchSize As Long) As Long
    On Error GoTo ErrorHandler
    If batchSize <= 0 Then
        CalculateTotalBatches = 0
    ElseIf totalIDs <= 0 Then
        CalculateTotalBatches = 0
    Else
        ' Ceiling division for calculating total batches
        CalculateTotalBatches = Application.WorksheetFunction.Ceiling(totalIDs / batchSize, 1)
    End If
    Exit Function
    
ErrorHandler:
    Log "Error in CalculateTotalBatches: " & Err.Description & " - returning zero", "WARNING"
    CalculateTotalBatches = 0
End Function

' Loop through batches and trigger ProcessSingleBatch
Private Sub ExecuteBatchLoop(colUnique As Collection, batchSize As Long, totalBatches As Long)
    On Error GoTo ErrorHandler
    
    Dim batchIdx As Long, i As Long, k As Long, endIdx As Long
    
    ' Process each batch
    For i = 1 To colUnique.Count Step batchSize
        Dim tBatch As Double: tBatch = Timer
        batchIdx = batchIdx + 1
        
        Log "========== Batch " & batchIdx & "/" & totalBatches & " =========="
        
        ' Build batch dictionary
        Dim batchIDs As Object: Set batchIDs = CreateObject("Scripting.Dictionary")
        ' Pre-size dictionary for better performance
        batchIDs.CompareMode = vbTextCompare
        
        endIdx = Application.Min(i + batchSize - 1, colUnique.Count)
        
        ' Populate batchIDs with error handling
        For k = i To endIdx
            On Error Resume Next
            batchIDs.Add colUnique(k), Empty
            If Err.Number <> 0 Then
                Log "Failed to add ID to batch " & batchIdx & ": " & Err.Description, "WARNING"
                Err.Clear
            End If
            On Error GoTo ErrorHandler
        Next k
        
        If batchIDs.Count = 0 Then
            Log "Batch " & batchIdx & " has 0 IDs - SKIPPED", "WARNING"
            GoTo NextBatch
        End If
        
        Log "  Contains " & batchIDs.Count & " MI_NOs (IDs " & i & " to " & endIdx & ")"
        
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
    If Not batchIDs Is Nothing Then Set batchIDs = Nothing
    Log "ExecuteBatchLoop ERROR at Batch " & batchIdx & ": " & Err.Description & " (Err# " & Err.Number & ")", "ERROR"
    Err.Raise Err.Number, "ExecuteBatchLoop", Err.Description
End Sub

' ===========ProcessSingleBatch=============
Private Function ProcessSingleBatch(batchIDs As Object, bIdx As Long) As Boolean
    ' Initialize return value
    ProcessSingleBatch = False
    On Error GoTo ErrorHandler
    
    ' Create temporary workbook
    Dim wbTemp As Workbook, wsTemp As Worksheet
    CreateTemporaryWorkbook wbTemp, wsTemp
    If wbTemp Is Nothing Or wsTemp Is Nothing Then
        Log "    Failed to create temp workbook", "ERROR"
        Exit Function
    End If
    
    ' Filter data for current batch
    If Not FilterBatchData(wsTemp, batchIDs) Then
        Log "    Failed to filter batch data", "ERROR"
        GoTo CleanBatch
    End If
    
    Dim filteredRows As Long: filteredRows = GetLastRow(wsTemp)
    Log "    Filtered " & (filteredRows - 1) & " data rows"
    
    If filteredRows < 2 Then
        Log "    No matching data - SKIPPED", "WARNING"
        ProcessSingleBatch = True ' Consider this successful as it's a valid scenario
        GoTo CleanBatch
    End If
    
    ' Refresh header map for temp worksheet
    RefreshHeaderMap wsTemp
    
    ' Execute batch helpers
    ExecuteBatchHelpers wsTemp
    
    ' Save batch output
    ProcessSingleBatch = SaveBatchOutput(wbTemp, bIdx)
    
CleanBatch:
    On Error Resume Next
    If Not wbTemp Is Nothing Then wbTemp.Close False
    Set wbTemp = Nothing
    Set wsTemp = Nothing
    On Error GoTo 0
    Exit Function
    
ErrorHandler:
    Log "BATCH ERROR " & bIdx & ": " & Err.Description & " (Err# " & Err.Number & ")", "ERROR"
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
    Log "    Error creating temp workbook: " & Err.Description & " (Err# " & Err.Number & ")", "ERROR"
    ' Ensure workbook is closed if it was partially created
    If Not wbTemp Is Nothing Then
        wbTemp.Close SaveChanges:=False
        Set wbTemp = Nothing
    End If
    Set wsTemp = Nothing
End Sub

' Call FilterDataToSheet for current batch
Private Function FilterBatchData(wsTemp As Worksheet, batchIDs As Object) As Boolean
    On Error GoTo FilterErr
    
    FilterBatchData = False
    
    FilterDataToSheet g_SourceWS, wsTemp, batchIDs
    FilterBatchData = True
    
    Exit Function
    
FilterErr:
    Log "FILTER ERROR: " & Err.Description & " (Err# " & Err.Number & ")", "ERROR"
End Function

' Execute all batch helpers
Private Sub ExecuteBatchHelpers(ws As Worksheet)
    On Error GoTo ErrorHandler
    
    RunAllHelpers ws
    Exit Sub
    
ErrorHandler:
    Log "HELPER EXECUTION ERROR: " & Err.Description & " (Err# " & Err.Number & ")", "ERROR"
End Sub

' Save batch output to file
Private Function SaveBatchOutput(wbTemp As Workbook, bIdx As Long) As Boolean
    On Error GoTo ErrorHandler
    
    SaveBatchOutput = False
    
    Dim fName As String
    fName = OUT_PATH & g_FSO.GetBaseName(g_SourceWB.Name) & _
            "_Batch_" & Format(bIdx, "000") & "_" & g_ProcessID & ".csv"
    
    Application.DisplayAlerts = False
    wbTemp.SaveAs fName, xlCSV
    Application.DisplayAlerts = True
    
    Log "    Saved: " & g_FSO.GetFileName(fName)
    SaveBatchOutput = True
    
    Exit Function
    
ErrorHandler:
    ' Ensure DisplayAlerts is re-enabled even if error occurs
    Application.DisplayAlerts = True
    Log "SAVE ERROR: " & Err.Description & " (Err# " & Err.Number & ")", "ERROR"
End Function

' ==============================================================================
' HELPER SYSTEM FUNCTIONS
' ==============================================================================

' ===========RunAllHelpers=============
Private Sub RunAllHelpers(ws As Worksheet)
    On Error GoTo ErrorHandler
    
    ExecuteHelper ws, "Helper_01_SubGroupMapping"
    ExecuteHelper ws, "Helper_02_Source"
    ExecuteHelper ws, "Helper_03_NB_and_RI"
    ExecuteHelper ws, "Helper_04_DefaultRate"
    ExecuteHelper ws, "Helper_05_LTVPropValue"
    ExecuteHelper ws, "Helper_06_AssumptionValuation"
    ExecuteHelper ws, "Helper_07_Policies_Date"
    ExecuteHelper ws, "Helper_08_ScenarioMerge"
    ExecuteHelper ws, "Helper_09_ReinsurerMapping"
    ExecuteHelper ws, "Helper_10_AddGMR_MTHLY"
    ExecuteHelper ws, "Helper_11_AddMaturity_Term"
    
    ExecuteHelper ws, "Helper_13_Projection_Expand"
    RefreshHeaderMap ws
    
    ExecuteHelper ws, "Helper_14_AddRemaining_Term"
    ExecuteHelper ws, "Helper_15_YrMth_Indicator"
    ExecuteHelper ws, "Helper_16_DefaultPattern"
    ExecuteHelper ws, "Helper_17_ClaimRate"
    ExecuteHelper ws, "Helper_18_CalculateOMCR"
    ExecuteHelper ws, "Helper_20_OMV_and_PctOMV_Combined"
    ExecuteHelper ws, "Helper_21_CPR_SMM_Combined"
    ExecuteHelper ws, "Helper_23_AcquisitionExpense"
    ExecuteHelper ws, "Helper_24_PolicyForceAndOPB_Combined"
    RefreshHeaderMap ws
    
    ExecuteHelper ws, "Helper_24A_Duration_Plus_and_Min"
    ExecuteHelper ws, "Helper_24B_Next_Policy_In_Force"
    ExecuteHelper ws, "Helper_24C_PreviousValues"
    RefreshHeaderMap ws
    
    ExecuteHelper ws, "Helper_25_FixedAssumedSeverity"
    ExecuteHelper ws, "Helper_26_InflationFactor"
    ExecuteHelper ws, "Helper_27_MaintenanceExpense"
    ExecuteHelper ws, "Helper_28_Commission_and_Commission_Recovery"
    ExecuteHelper ws, "Helper_29_RI_Policies_Count"

    ExecuteHelper ws, "Helper_30_RI_Premium"
    ExecuteHelper ws, "Helper_31_RiskInForce_and_DefaultClaimOutgo"
    ExecuteHelper ws, "Helper_32_RI_Collateral"
    ExecuteHelper ws, "Helper_33_RI_NPR"
    
    Exit Sub
    
ErrorHandler:
    Log "RunAllHelpers ERROR: " & Err.Description & " (Err# " & Err.Number & ")", "ERROR"
End Sub

Private Sub ExecuteHelper(ws As Worksheet, procName As String)
    On Error GoTo ErrorHandler
    
    Dim t As Double: t = Timer
    Application.Run procName, ws
    Log "      " & procName & ": " & Format(Timer - t, "0.000") & "s", "SUCCESS"
    Exit Sub
    
ErrorHandler:
    Log "      " & procName & " ERROR(" & Err.Number & "): " & Err.Description, "ERROR"
    Err.Clear
End Sub

' ==============================================================================
' FILTERING ENGINE
' ==============================================================================

Private Sub BuildRowIndex(ws As Worksheet)
    On Error GoTo IndexErr
    
    Set g_RowIndex = CreateObject("Scripting.Dictionary")
    ' Pre-size dictionary for better performance
    g_RowIndex.CompareMode = vbTextCompare
    
    Dim lr As Long: lr = GetLastRow(ws)
    If lr < 2 Then Exit Sub
    
    Dim vMI As Variant
    On Error Resume Next
    vMI = ws.Range(ws.Cells(2, 1), ws.Cells(lr, 1)).Value2
    If Err.Number <> 0 Then
        Log "  BuildRowIndex: Failed to read MI_NO column - " & Err.Description, "ERROR"
        Err.Clear
        Exit Sub
    End If
    On Error GoTo IndexErr
    
    vMI = SafeArray2D(vMI)
    Dim rowCount As Long: rowCount = UBound(vMI, 1)
    
    Dim i As Long, miNo As String, rowCol As Collection
    
    For i = 1 To rowCount
        miNo = Trim$(CStr(vMI(i, 1)))
        
        If Len(miNo) > 0 Then
            If Not g_RowIndex.Exists(miNo) Then
                Set rowCol = New Collection
                g_RowIndex.Add rowCol, miNo
            Else
                Set rowCol = g_RowIndex(miNo)
            End If
            rowCol.Add i + 1
        End If
    Next i
    
    Log "  Indexed " & g_RowIndex.Count & " unique MI_NOs"
    
    Erase vMI
    Set rowCol = Nothing
    Exit Sub
    
IndexErr:
    Log "BuildRowIndex ERROR: " & Err.Description & " (Err# " & Err.Number & ")", "ERROR"
    Set g_RowIndex = CreateObject("Scripting.Dictionary")
End Sub

Private Sub FilterDataToSheet(srcWS As Worksheet, destWS As Worksheet, validIDs As Object)
    On Error GoTo FilterErr
    
    Dim lr As Long, lc As Long
    Dim vSrc As Variant, vRes As Variant
    Dim srcRows As Long, srcCols As Long
    Dim outR As Long, r As Long, c As Long
    Dim miNo As Variant, rowCol As Collection, rowNum As Variant
    
    If srcWS Is Nothing Then 
        Log "FilterDataToSheet: Source worksheet is Nothing", "ERROR"
        Err.Raise 91, , "Source worksheet is Nothing"
    End If
    
    If destWS Is Nothing Then 
        Log "FilterDataToSheet: Destination worksheet is Nothing", "ERROR"
        Err.Raise 91, , "Destination worksheet is Nothing"
    End If
    
    lr = GetLastRow(srcWS)
    If lr < 1 Then
        Log "      No data in source", "WARNING"
        Exit Sub
    End If
    
    lc = GetLastColumn(srcWS)
    
    ' Copy header
    srcWS.Rows(1).Copy destWS.Rows(1)
    
    If lr < 2 Then Exit Sub
    
    ' Read source
    On Error Resume Next
    vSrc = srcWS.Range("A1", srcWS.Cells(lr, lc)).Value2
    If Err.Number <> 0 Then
        Log "      ERROR reading source: " & Err.Description, "ERROR"
        Err.Raise 51, "FilterDataToSheet", "Error reading source data: " & Err.Description
    End If
    On Error GoTo FilterErr
    
    vSrc = SafeArray2D(vSrc)
    srcRows = UBound(vSrc, 1)
    srcCols = UBound(vSrc, 2)
    
    ReDim vRes(1 To srcRows, 1 To srcCols)
    
    ' Use index if available
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
        ' Fallback: Full scan
        For r = 2 To srcRows
            On Error Resume Next
            miNo = Trim$(CStr(vSrc(r, 1)))
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
    
    ' Clean up
    Erase vSrc, vRes
    Set rowCol = Nothing
    Exit Sub
    
FilterErr:
    ' Ensure cleanup happens even in error cases
    On Error Resume Next
    Erase vSrc, vRes
    Set rowCol = Nothing
    On Error GoTo 0
    
    Log "FilterDataToSheet ERROR: " & Err.Description & " (Err# " & Err.Number & ")", "ERROR"
    Err.Raise Err.Number, "FilterDataToSheet", Err.Description
End Sub

Private Function GetDistinctIDs(ws As Worksheet) As Collection
    On Error GoTo IDErr
    
    Set GetDistinctIDs = New Collection
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    ' Pre-size dictionary for better performance
    d.CompareMode = vbTextCompare
    
    Dim lr As Long
    lr = GetLastRow(ws)
    If lr < 2 Then 
        Set d = Nothing
        Exit Function
    End If
    
    ' Get data range
    Dim v As Variant
    On Error Resume Next
    v = ws.Range(ws.Cells(2, 1), ws.Cells(lr, 1)).Value2
    If Err.Number <> 0 Then
        Log "GetDistinctIDs: Failed to read data range - " & Err.Description, "ERROR"
        Err.Clear
        Set d = Nothing
        Exit Function
    End If
    On Error GoTo IDErr
    
    v = SafeArray2D(v)
    Dim rowCount As Long: rowCount = UBound(v, 1)
    
    Dim i As Long
    Dim s As String
    
    For i = 1 To rowCount
        s = Trim$(CStr(v(i, 1)))
        If Len(s) > 0 And Not d.Exists(s) Then
            d(s) = Empty
            On Error Resume Next
            GetDistinctIDs.Add s
            If Err.Number <> 0 Then
                Log "Failed to add item to collection: " & Err.Description, "WARNING"
                Err.Clear
            End If
            On Error GoTo IDErr
        End If
    Next i
    
    ' Clean up resources
    Set d = Nothing
    Erase v
    Exit Function
    
IDErr:
    Set d = Nothing
    Log "GetDistinctIDs ERROR: " & Err.Description & " (Err# " & Err.Number & ")", "ERROR"
    Set GetDistinctIDs = New Collection
End Function

' ==============================================================================
' COLUMN MANAGEMENT FUNCTIONS
' ==============================================================================

' ================= CONSOLIDATED AddColumn - HANDLES ALL SITUATIONS =================
Public Sub AddColumn(ws As Worksheet, colName As String)
    On Error GoTo ErrHandler
    
    ' Validate inputs
    If ws Is Nothing Then
        Log "AddColumn: Worksheet parameter is Nothing", "ERROR"
        Exit Sub
    End If
    
    If Len(Trim$(colName)) = 0 Then
        Log "AddColumn: Empty column name provided", "WARNING"
        Exit Sub
    End If
    
    Dim ucName As String
    ucName = UCase$(Trim$(Replace(colName, Chr(160), " "))) ' Normalize
    
    ' Step 1: Refresh header map to avoid stale state
    On Error Resume Next
    RefreshHeaderMap ws
    If Err.Number <> 0 Then
        Log "AddColumn: ERROR in RefreshHeaderMap: " & Err.Description, "WARNING"
        Err.Clear
    End If
    On Error GoTo ErrHandler
    
    ' Step 2: Check if column already exists
    If g_HeaderMap.Exists(ucName) Then
        ' Column exists, verify it's valid
        On Error Resume Next
        Dim existingCol As Long: existingCol = g_HeaderMap(ucName)
        Dim existingHeader As String: existingHeader = ws.Cells(1, existingCol).Value
        If Err.Number = 0 And UCase$(Trim$(existingHeader)) = ucName Then
            ' Valid existing column
            Exit Sub
        Else
            ' Stale map entry, remove it
            Log "AddColumn: Stale map entry for " & colName & ", refreshing", "DEBUG"
            g_HeaderMap.Remove ucName
            Err.Clear
        End If
        On Error GoTo ErrHandler
    End If
    
    ' Step 3: Determine last column position SAFELY
    Dim lc As Long: lc = 0
    
    ' Method 1: Use HeaderMap (fastest if available)
    On Error Resume Next
    If g_HeaderMap.Count > 0 Then
        ' Find max column index more efficiently
        Dim vals As Variant: vals = g_HeaderMap.Items
        Dim maxCol As Long: maxCol = 0
        Dim i As Long
        For i = 0 To UBound(vals)
            Dim colIdx As Long: colIdx = CLng(vals(i))
            If Err.Number = 0 And colIdx > maxCol Then maxCol = colIdx
        Next i
        If maxCol > 0 Then lc = maxCol
    End If
    Err.Clear
    On Error GoTo ErrHandler
    
    ' Method 2: Use UsedRange (fallback)
    If lc = 0 Then
        On Error Resume Next
        lc = ws.UsedRange.Columns.Count
        If Err.Number <> 0 Or lc = 0 Then
            Err.Clear
            ' Method 3: Scan row 1 (last resort)
            lc = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
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
        If UCase$(Trim$(CStr(headerVal))) = ucName Then
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
        Log "AddColumn: Cannot add column " & colName & " - Excel column limit reached", "ERROR"
        Exit Sub
    End If
    
    ' Step 5: Add header with error handling
    On Error Resume Next
    ws.Cells(1, targetCol).Value = colName
    If Err.Number <> 0 Then
        Log "AddColumn: ERROR writing header to column " & targetCol & ": " & Err.Description, "ERROR"
        Err.Clear
        Exit Sub
    End If
    On Error GoTo ErrHandler
    
    ' Step 6: Update header map
    On Error Resume Next
    g_HeaderMap(ucName) = targetCol
    If Err.Number <> 0 Then
        Log "AddColumn: ERROR updating HeaderMap: " & Err.Description, "WARNING"
        Err.Clear
    End If
    On Error GoTo ErrHandler
    
    ' Step 7: Verify addition
    On Error Resume Next
    Dim verifyHeader As String: verifyHeader = ws.Cells(1, targetCol).Value
    If Err.Number = 0 And UCase$(Trim$(verifyHeader)) = ucName Then
        ' Success - no log needed for normal operation
    Else
        Log "AddColumn: WARNING - Column " & colName & " may not have been added correctly", "WARNING"
    End If
    Err.Clear
    On Error GoTo ErrHandler
    
    Exit Sub

ErrHandler:
    Log "AddColumn: CRITICAL ERROR for column " & colName & ": " & Err.Description & " (Err#" & Err.Number & ")", "ERROR"
    
    ' Attempt recovery: try direct write to next available column
    On Error Resume Next
    Dim recoveryCol As Long: recoveryCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column + 1
    If recoveryCol > 0 And recoveryCol <= 16384 Then
        ws.Cells(1, recoveryCol).Value = colName
        If Err.Number = 0 Then
            g_HeaderMap(ucName) = recoveryCol
            Log "AddColumn: Recovery successful - added " & colName & " at column " & recoveryCol, "WARNING"
        End If
    End If
    Err.Clear
    On Error GoTo 0
End Sub

Public Function GetColumnIndex(ws As Worksheet, colName As String) As Long
    On Error GoTo ErrorHandler
    
    ' Validate inputs
    If ws Is Nothing Then
        Log "GetColumnIndex: Worksheet parameter is Nothing", "ERROR"
        GetColumnIndex = 0
        Exit Function
    End If
    
    If Len(Trim$(colName)) = 0 Then
        GetColumnIndex = 0
        Exit Function
    End If
    
    Dim ucName As String: ucName = UCase$(Trim$(colName))
    If Not g_HeaderMap Is Nothing Then
        If g_HeaderMap.Exists(ucName) Then
            GetColumnIndex = g_HeaderMap(ucName)
        Else
            GetColumnIndex = 0
        End If
    Else
        GetColumnIndex = 0
    End If
    Exit Function
    
ErrorHandler:
    Log "GetColumnIndex ERROR: " & Err.Description & " (Err# " & Err.Number & ") - returning 0", "ERROR"
    GetColumnIndex = 0
End Function

Private Sub RefreshHeaderMap(ws As Worksheet)
    On Error GoTo RefreshErr
    
    ' Validate input
    If ws Is Nothing Then
        Log "RefreshHeaderMap: Worksheet parameter is Nothing", "ERROR"
        Exit Sub
    End If
    
    g_HeaderMap.RemoveAll
    
    Dim vHead As Variant
    On Error Resume Next
    vHead = ws.Rows(1).Value2
    If Err.Number <> 0 Then
        Log "RefreshHeaderMap: Failed to read headers - " & Err.Description, "ERROR"
        Err.Clear
        Exit Sub
    End If
    On Error GoTo RefreshErr
    
    vHead = SafeArray2D(vHead)
    Dim colCount As Long: colCount = UBound(vHead, 2)
    Dim c As Long, colName As String
    
    For c = 1 To colCount
        If Not (IsError(vHead(1, c)) Or IsEmpty(vHead(1, c))) Then
            colName = Trim$(CStr(vHead(1, c)))
            If Len(colName) > 0 Then 
                On Error Resume Next
                g_HeaderMap(UCase$(colName)) = c
                If Err.Number <> 0 Then
                    Log "Failed to add column to header map: " & Err.Description, "WARNING"
                    Err.Clear
                End If
                On Error GoTo RefreshErr
            End If
        End If
    Next c
    
    Erase vHead
    Exit Sub
    
RefreshErr:
    If Not IsEmpty(vHead) Then Erase vHead
    Log "RefreshHeaderMap ERROR: " & Err.Description & " (Err# " & Err.Number & ")", "ERROR"
End Sub

' ==============================================================================
' TYPE-SAFE CONVERSION FUNCTIONS
' ==============================================================================

' ======== CONSOLIDATED NZ FUNCTIONS - EXPLICIT PARAMETERS ========
Public Function NzLong(data As Variant, row As Long, col As Long) As Long
    On Error GoTo ErrorHandler
    
    ' Validate inputs
    If Not IsArray(data) Then
        NzLong = 0
        Exit Function
    End If
    
    Dim v As Variant
    
    On Error Resume Next
    Dim uRow As Long, uCol As Long, lRow As Long, lCol As Long
    uRow = UBound(data, 1)
    uCol = UBound(data, 2)
    lRow = LBound(data, 1)
    lCol = LBound(data, 2)
    On Error GoTo ErrorHandler
    
    If (lRow > 0 And lCol > 0 And row >= lRow And col >= lCol And row <= uRow And col <= uCol) Or _
       (lRow <= 0 And lCol <= 0 And row >= lRow And col >= lCol And row <= uRow And col <= uCol) Then
        On Error Resume Next
        v = data(row, col)
        If Err.Number = 0 Then
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
        On Error GoTo ErrorHandler
    Else
        NzLong = 0
    End If
    
    Exit Function
    
ErrorHandler:
    NzLong = 0
End Function

' ======== CONSOLIDATED NZ FUNCTIONS - EXPLICIT PARAMETERS ========
Public Function NzDouble(data As Variant, r As Long, c As Long) As Double
    On Error GoTo ErrorHandler
    
    ' Validate inputs
    If Not IsArray(data) Then
        NzDouble = 0#
        Exit Function
    End If
    
    On Error Resume Next
    Dim v As Variant: v = data(r, c)
    If IsEmpty(v) Or IsNull(v) Or IsError(v) Then
        NzDouble = 0#
    ElseIf IsNumeric(v) Then
        NzDouble = CDbl(v)
        If Err.Number <> 0 Then
            NzDouble = 0#
            Err.Clear
        End If
    Else
        NzDouble = 0#
    End If
    On Error GoTo ErrorHandler
    
    Exit Function
    
ErrorHandler:
    NzDouble = 0#
End Function

' ======== CONSOLIDATED NZ FUNCTIONS - EXPLICIT PARAMETERS ========
Public Function NzString(data As Variant, r As Long, c As Long) As String
    On Error GoTo ErrorHandler
    
    ' Validate inputs
    If Not IsArray(data) Then
        NzString = ""
        Exit Function
    End If
    
    Dim v As Variant
    On Error Resume Next
    v = data(r, c)
    On Error GoTo ErrorHandler
    
    If IsEmpty(v) Or IsNull(v) Or IsError(v) Then
        NzString = ""
    Else
        NzString = Trim$(CStr(v))
    End If
    
    Exit Function
    
ErrorHandler:
    NzString = ""
End Function

' ==============================================================================
' ======== CONSOLIDATED DATE FUNCTIONS - HANDLES YYYYMMDD FORMAT ===========
Public Function ParseDate(ByVal v As Variant, ByRef outDate As Date) As Boolean
    On Error GoTo ErrorHandler
    
    ' Handle Empty/Null/Error
    If IsEmpty(v) Or IsNull(v) Or IsError(v) Then 
        ParseDate = False
        Exit Function
    End If
    
    ' Numeric branch
    If IsNumeric(v) Then
        Dim dblVal As Double: dblVal = CDbl(v)
        
        ' YYYYMMDD format
        If dblVal >= 19000101 And dblVal <= 99991231 Then
            Dim y As Long, m As Long, d As Long
            y = Int(dblVal / 10000)
            m = Int((dblVal - y * 10000) / 100)
            d = dblVal - y * 10000 - m * 100
            ' Using DateSerial's built-in date validation rather than manual checking
            outDate = DateSerial(y, m, d)
            ' Check that the constructed date can correctly convert back to YYYYMMDD format to confirm validity
            If CLng(Year(outDate)) * 10000 + CLng(Month(outDate)) * 100 + CLng(Day(outDate)) = dblVal Then
                ParseDate = True
            End If
        ' Excel serial
        ElseIf dblVal > 0 And dblVal < 2958466 Then
            outDate = CDate(dblVal)
            ParseDate = True
        End If
    
    ' Text branch (CSV often loads dates as text)
    ElseIf VarType(v) = vbString Then
        If IsDate(v) Then
            outDate = CDate(v)
            ParseDate = True
        End If
    End If
    
    Exit Function
    
ErrorHandler:
    ParseDate = False
    Err.Clear
End Function

' =================== SAFE ARRAY WRAPPER ===================
Private Function SafeArray2D(inputValue As Variant) As Variant
    On Error GoTo ErrorHandler
    
    ' Handle non-array values
    If Not IsArray(inputValue) Then
        Dim singleValueArray(1 To 1, 1 To 1) As Variant
        singleValueArray(1, 1) = inputValue
        SafeArray2D = singleValueArray
        Exit Function
    End If
    
    ' Try to use as 2D array directly
    On Error Resume Next
    Dim testRows As Long: testRows = UBound(inputValue, 1)
    Dim testCols As Long: testCols = UBound(inputValue, 2)
    On Error GoTo ErrorHandler
    
    If testRows >= 1 And testCols >= 1 Then
        ' Valid 2D array
        SafeArray2D = inputValue
    Else
        ' Convert 1D array to 2D
        On Error Resume Next
        Dim arrayBound As Long: arrayBound = UBound(inputValue)
        Dim arrayLBound As Long: arrayLBound = LBound(inputValue)
        On Error GoTo ErrorHandler
        
        ' Create 2D array with same number of elements
        Dim converted2D() As Variant
        If arrayBound >= arrayLBound Then
            ' Have valid elements
            ReDim converted2D(1 To 1, arrayLBound To arrayBound)
            Dim i As Long
            For i = arrayLBound To arrayBound
                converted2D(1, i) = inputValue(i)
            Next i
        Else
            ' Empty or invalid array, create minimal 2D array
            ReDim converted2D(1 To 1, 1 To 1)
        End If
        SafeArray2D = converted2D
    End If
    
    Exit Function
    
ErrorHandler:
    ' Return empty 2D array on error
    Dim emptyArray(1 To 1, 1 To 1) As Variant
    SafeArray2D = emptyArray
    Log "SafeArray2D ERROR: " & Err.Description & " - returning empty array", "ERROR"
End Function

' =================== CONSOLIDATED PROCESSING ===================
Private Sub ProcessCommissionData(ws As Worksheet, lr As Long, _
                              cOpt As Long, cAmt As Long, cRate As Long, cRIPct As Long, _
                              cYr As Long, cMon As Long, cDur As Long, cRI As Long, _
                              cComm As Long, cCommBOP As Long, cRIComm As Long, cRICommBOP As Long, _
                              useChunking As Boolean, is64Bit As Boolean)
    
    On Error GoTo ErrorHandler
    
    ' Validate inputs
    If ws Is Nothing Then
        Log "ProcessCommissionData: Worksheet parameter is Nothing", "ERROR"
        Exit Sub
    End If
    
    ' If not using chunking, process all at once
    If Not useChunking Then
        ProcessSinglePass ws, lr, cOpt, cAmt, cRate, cRIPct, cYr, cMon, cDur, cRI, cComm, cCommBOP, cRIComm, cRICommBOP
        Exit Sub
    End If
    
    ' Otherwise, process in chunks
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
    Dim hasAmountError As Boolean, hasRateError As Boolean, hasRIPctError As Boolean
    Dim hasCriticalError As Boolean
    Dim convResult As Variant
    
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
            hasCriticalError = False
            
            If VarType(v) = vbString Then
                optType = UCase$(Trim$(CStr(v)))
                
                If optType = "SI" Then
                    ' Initialize variables
                    amt = 0: rate = 0: riPct = 0: yr = 0: mon = 0: dur = 0: isRI = False
                    hasAmountError = False: hasRateError = False: hasRIPctError = False
                    
                    ' Process Amount
                    v = vAmt(r, 1)
                    If IsNumeric(v) And Not IsEmpty(v) Then
                        On Error Resume Next
                        If Not TryConvertVariant(v, convResult, "Double") Then
                            hasAmountError = True
                            hasCriticalError = True
                        Else
                            amt = convResult
                        End If
                        Err.Clear
                        On Error GoTo ErrorHandler
                    Else
                        hasAmountError = True
                        hasCriticalError = True
                    End If
                    
                    ' Process Rate
                    v = vRate(r, 1)
                    If IsNumeric(v) And Not IsEmpty(v) Then
                        On Error Resume Next
                        If Not TryConvertVariant(v, convResult, "Double") Then
                            hasRateError = True
                            hasCriticalError = True
                        Else
                            rate = convResult
                        End If
                        Err.Clear
                        On Error GoTo ErrorHandler
                    Else
                        hasRateError = True
                        hasCriticalError = True
                    End If
                    
                    ' Process RI Percentage
                    v = vRIPct(r, 1)
                    If IsNumeric(v) And Not IsEmpty(v) Then
                        On Error Resume Next
                        If TryConvertVariant(v, convResult, "Double") Then
                            riPct = convResult * 0.01
                        Else
                            hasRIPctError = True
                            hasCriticalError = True
                        End If
                        Err.Clear
                        On Error GoTo ErrorHandler
                    Else
                        hasRIPctError = True
                        hasCriticalError = True
                    End If
                    
                    ' Count as error if any critical fields have errors
                    If hasCriticalError Then
                        totalErrors = totalErrors + 1
                    End If
                    
                    ' Process Year
                    v = vYr(r, 1)
                    If IsNumeric(v) And Not IsEmpty(v) Then
                        On Error Resume Next
                        If Not TryConvertVariant(v, convResult, "Long") Then
                            yr = 0
                        Else
                            yr = convResult
                        End If
                        Err.Clear
                        On Error GoTo ErrorHandler
                    End If
                    
                    ' Process Month
                    v = vMon(r, 1)
                    If IsNumeric(v) And Not IsEmpty(v) Then
                        On Error Resume Next
                        If Not TryConvertVariant(v, convResult, "Long") Then
                            mon = 0
                        Else
                            mon = convResult
                        End If
                        Err.Clear
                        On Error GoTo ErrorHandler
                    End If
                    
                    ' Process Duration
                    v = vDur(r, 1)
                    If IsNumeric(v) And Not IsEmpty(v) Then
                        On Error Resume Next
                        If Not TryConvertVariant(v, convResult, "Long") Then
                            dur = 0
                        Else
                            dur = convResult
                        End If
                        Err.Clear
                        On Error GoTo ErrorHandler
                    End If
                    
                    ' Process RI flag
                    v = vRI(r, 1)
                    If VarType(v) = vbString Then
                        isRI = (UCase$(Trim$(CStr(v))) = "Y")
                    ElseIf IsNumeric(v) Then
                        On Error Resume Next
                        If TryConvertVariant(v, convResult, "Long") Then
                            isRI = (convResult = 1)
                        End If
                        Err.Clear
                        On Error GoTo ErrorHandler
                    End If
                    
                    ' Calculate commission only once if needed
                    If (yr = 1 And mon = 1) Or (dur = 1) Then
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
            End If
        Next r
        
        ws.Cells(chunkStart, cComm).Resize(rows, 4).Value2 = vOut
        
        Erase vOpt, vAmt, vRate, vRIPct, vYr, vMon, vDur, vRI, vOut
        
        chunkStart = chunkEnd + 1
    Loop
    
    If totalErrors > 0 Then Log "Commission processing: " & totalErrors & " conversion errors", "WARNING"
    Exit Sub
    
ErrorHandler:
    Log "ProcessCommissionData ERROR: " & Err.Description & " (Err# " & Err.Number & ")", "ERROR"
    Erase vOpt, vAmt, vRate, vRIPct, vYr, vMon, vDur, vRI, vOut
End Sub

' =================== PROCESS SINGLE PASS ===================
Private Sub ProcessSinglePass(ws As Worksheet, lr As Long, _
                              cOpt As Long, cAmt As Long, cRate As Long, cRIPct As Long, _
                              cYr As Long, cMon As Long, cDur As Long, cRI As Long, _
                              cComm As Long, cCommBOP As Long, cRIComm As Long, cRICommBOP As Long)
    
    On Error GoTo ErrorHandler
    
    ' Validate inputs
    If ws Is Nothing Then
        Log "ProcessSinglePass: Worksheet parameter is Nothing", "ERROR"
        Exit Sub
    End If
    
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
    Dim convResult As Variant
    
    For r = 1 To rows
        v = vOpt(r, 1)
        If VarType(v) = vbString Then
            optType = UCase$(Trim$(CStr(v)))
            
            If optType = "SI" Then
                ' Initialize variables
                amt = 0: rate = 0: riPct = 0: yr = 0: mon = 0: dur = 0: isRI = False
                
                ' Process Amount
                v = vAmt(r, 1)
                If IsNumeric(v) And Not IsEmpty(v) Then
                    If Not TryConvertVariant(v, convResult, "Double") Then 
                        errCount = errCount + 1
                    Else
                        amt = convResult
                    End If
                Else
                    errCount = errCount + 1
                End If
                
                ' Process Rate
                v = vRate(r, 1)
                If IsNumeric(v) And Not IsEmpty(v) Then
                    If Not TryConvertVariant(v, convResult, "Double") Then 
                        errCount = errCount + 1
                    Else
                        rate = convResult
                    End If
                Else
                    errCount = errCount + 1
                End If
                
                ' Process RI Percentage
                v = vRIPct(r, 1)
                If IsNumeric(v) And Not IsEmpty(v) Then
                    If TryConvertVariant(v, convResult, "Double") Then
                        riPct = convResult * 0.01
                    Else
                        errCount = errCount + 1
                    End If
                Else
                    errCount = errCount + 1
                End If
                
                ' Process Year
                v = vYr(r, 1)
                If IsNumeric(v) And Not IsEmpty(v) Then
                    If Not TryConvertVariant(v, convResult, "Long") Then 
                        yr = 0
                    Else
                        yr = convResult
                    End If
                End If
                
                ' Process Month
                v = vMon(r, 1)
                If IsNumeric(v) And Not IsEmpty(v) Then
                    If Not TryConvertVariant(v, convResult, "Long") Then 
                        mon = 0
                    Else
                        mon = convResult
                    End If
                End If
                
                ' Process Duration
                v = vDur(r, 1)
                If IsNumeric(v) And Not IsEmpty(v) Then
                    If Not TryConvertVariant(v, convResult, "Long") Then 
                        dur = 0
                    Else
                        dur = convResult
                    End If
                End If
                
                ' Process RI flag
                v = vRI(r, 1)
                If VarType(v) = vbString Then
                    isRI = (UCase$(Trim$(CStr(v))) = "Y")
                ElseIf IsNumeric(v) Then
                    Dim tempLong As Long
                    If TryConvertVariant(v, convResult, "Long") Then
                        isRI = (convResult = 1)
                    End If
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
    
    If errCount > 0 Then Log "Commission processing: " & errCount & " conversion errors", "WARNING"
    Exit Sub
    
ErrorHandler:
    Log "ProcessSinglePass ERROR: " & Err.Description & " (Err# " & Err.Number & ")", "ERROR"
    Erase vOpt, vAmt, vRate, vRIPct, vYr, vMon, vDur, vRI, vOut
End Sub

' Helper function to safely convert variant to double or long
Private Function TryConvertVariant(value As Variant, ByRef result As Variant, targetType As String) As Boolean
    On Error GoTo ErrorHandler
    
    Select Case targetType
        Case "Double"
            result = CDbl(value)
            TryConvertVariant = True
        Case "Long"
            result = CLng(value)
            TryConvertVariant = True
        Case Else
            TryConvertVariant = False
    End Select
    
    Exit Function
    
ErrorHandler:
    TryConvertVariant = False
    Err.Clear
End Function

' ===========LoadDefaultPatterns - LOAD PATTERN LOOKUP=============
Private Sub LoadDefaultPatterns()
    On Error GoTo LoadPatErr
    
    ' Find pattern sheet
    Dim patternWS As Worksheet
    Set patternWS = Nothing
    On Error Resume Next
    Set patternWS = ThisWorkbook.Sheets("DEFAULT_PATTERN")
    On Error GoTo 0 ' Reset error handling
    
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
    On Error GoTo 0 ' Reset error handling
    
    vPat = SafeArray2D(vPat)
    
    ' Build pattern dictionary
    Set g_patternDict = CreateObject("Scripting.Dictionary")
    ' Pre-size dictionary for better performance
    g_patternDict.CompareMode = vbTextCompare
    
    Dim r As Long, key As String
    Dim rowCount As Long: rowCount = UBound(vPat, 1)
    
    For r = 2 To rowCount ' Skip header
        ' Check if key column is not empty
        If Not IsEmpty(vPat(r, 1)) Then
            key = Trim$(CStr(vPat(r, 1)))
            If Len(key) > 0 Then
                ' Store pattern value
                On Error Resume Next
                g_patternDict(key) = vPat(r, 2)
                If Err.Number <> 0 Then
                    Log "Failed to add pattern to dictionary: " & Err.Description, "WARNING"
                    Err.Clear
                End If
                On Error GoTo 0 ' Reset error handling
            End If
        End If
    Next r
    
    If g_patternDict.Count > 0 Then
        Log "LoadDefaultPatterns: Loaded " & g_patternDict.Count & " patterns"
    Else
        Log "LoadDefaultPatterns: No patterns found", "WARNING"
    End If
    
    ' Cleanup
    Set patternWS = Nothing
    Erase vPat
    Exit Sub
    
LoadPatErr:
    Log "LoadDefaultPatterns ERROR: " & Err.Description & " (Err# " & Err.Number & ")", "ERROR"
End Sub

' ==============================================================================
' END OF MAIN MODULE
' ==============================================================================