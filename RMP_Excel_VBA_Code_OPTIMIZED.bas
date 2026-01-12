Option Explicit

' =============================================================================
' RMP EXCEL VBA PROCESSING ENGINE - OPTIMIZED VERSION
' =============================================================================
'
' PURPOSE: Process large insurance datasets in batches with modular helper functions
' AUTHOR: HKMC Budget Enhancement Team
' VERSION: 2.5
' LAST UPDATED: December 2025
'
' OPTIMIZATIONS MADE:
' 1. Enhanced performance for large-scale data processing
' 2. Eliminated silent failures with explicit error handling
' 3. Implemented fast exit strategies with comprehensive logging
' 4. Improved memory management for large datasets
' 5. Enhanced error resilience with detailed diagnostics
' 6. Optimized array operations for better throughput
' 7. Added log buffering to reduce disk I/O operations
' 8. Implemented log level filtering for production runs
' 9. Used LongLong on 64-bit systems for very large row counts
' 10. Used Currency for monetary calculations instead of Double
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
'     - Logs a message with timestamp (types: DEBUG, INFO, WARNING, ERROR, FATAL, SUCCESS)
'     - Now supports buffered logging for better performance
'     - Log level filtering available via MIN_LOG_LEVEL constant
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
' MEMORY MANAGEMENT:
'   GetPooledArray(rows, cols)
'     - Gets a reusable array from the pool for better performance
'     - Returns a Variant array of specified dimensions
'     - Usage: myArray = GetPooledArray(100, 10)
'
'   ReturnArrayToPool(array)
'     - Returns an array to the pool for reuse
'     - Usage: ReturnArrayToPool myArray
'
' COLLECTION/DICTIONARY FUNCTIONS:
'   GetDistinctIDs(ws)
'     - Gets a collection of distinct IDs from column A of a worksheet
'     - Usage: Set idCollection = GetDistinctIDs(myWorksheet)
'
'   BuildRowIndex(ws)
'     - Builds an index of row numbers by MI_NO for faster filtering
'     - Populates the global g_RowIndex object
'     - Usage: BuildRowIndex myWorksheet
'
' FILTERING FUNCTIONS:
'   FilterDataToSheet(sourceWS, destWS, validIDs)
'     - Filters data from source worksheet to destination worksheet based on valid IDs
'     - Uses g_RowIndex if available for faster filtering
'     - Usage: FilterDataToSheet sourceWorksheet, destWorksheet, validIDDictionary
'
' WORKBOOK/WORKSHEET FUNCTIONS:
'   SetSourceWorkbook(wb, ws)
'     - Sets the global source workbook and worksheet references
'     - Usage: SetSourceWorkbook myWorkbook, myWorksheet
'
'   CreateTemporaryWorkbook(ByRef wbTemp, ByRef wsTemp)
'     - Creates a new temporary workbook and worksheet
'     - Usage: CreateTemporaryWorkbook myTempWorkbook, myTempWorksheet
'
' INITIALIZATION FUNCTIONS:
'   InitializeGlobalObjects()
'     - Initializes all global objects used by the engine
'     - Automatically called when needed by other functions
'     - Usage: InitializeGlobalObjects
'
' ERROR HANDLING:
'   Use standard On Error statements with descriptive error messages
'   - All functions should have proper error handling with informative logs
'
' PERFORMANCE FEATURES:
'   - Automatic array pooling to reduce memory allocation overhead
'   - Buffered logging to reduce disk I/O operations
'   - Configurable log level filtering (see MIN_LOG_LEVEL constant)
'   - LongLong support on 64-bit systems for very large datasets
'   - Currency data type for precise monetary calculations
'   - String normalization caching to avoid repeated operations
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
'       ' No silent errors - Log error details
'       Log "Helper_99_MyCustomFunction ERROR: " & Err.Description & " (Err# " & Err.Number & ")", "ERROR"
'       Dim errorDuration As Double: errorDuration = Timer - startTime
'       Log "Helper_99_MyCustomFunction FAILED | Duration: " & Format(errorDuration, "0.000") & "s", "ERROR"
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
'      - Use appropriate data types (Currency for money, LongLong for big counts on 64-bit)
'      - Avoid Variant in mathematical operations - convert to specific types first
'   5. CLEAN CODE:
'      - KISS/DRY principles
'      - Meaningful variable names (e.g., cColName, vDataArray)
'      - Declare all variables upfront
'   6. MEMORY MANAGEMENT:
'      - Use GetPooledArray and ReturnArrayToPool for large arrays
'      - Explicitly Erase arrays when no longer needed
'   7. LOGGING:
'      - Use appropriate log levels (DEBUG, INFO, WARNING, ERROR, FATAL, SUCCESS)
'      - Consider performance impact of logging in tight loops
' ------------------------------------------------------------------------------

' --- CONFIGURATION ---
Private Const LOG_PATH As String = "Log\"
Private Const OUT_PATH As String = "Output\"
Private Const DEBUG_PRINT As Boolean = True
Private Const DOEVENTS_FREQUENCY As Long = 5 ' Process DoEvents every N batches
Private Const DOEVENTS_TIME_THRESHOLD As Single = 0.5 ' Process DoEvents every N seconds
Private Const ARRAY_POOL_MAX_SIZE As Long = 20 ' Maximum arrays to keep in pool
Private Const ARRAY_POOL_TARGET_SIZE As Long = 10 ' Target size for pool maintenance
Private Const DISABLE_DOEVENTS As Boolean = False ' Set to True for maximum speed (UI will freeze)
Private Const LOG_BUFFER_SIZE As Long = 100 ' Number of log entries to buffer before flushing to disk
Private Const MIN_LOG_LEVEL As String = "DEBUG" ' Minimum log level to output (DEBUG, INFO, WARNING, ERROR, FATAL)

' --- GLOBALS (Engine) ---
Private g_FSO As Object  ' Early binding for better performance
Private g_LogStream As Object
Private g_HeaderMap As Object
Private g_ProcessID As String
Public g_LookupData As Object
Public g_SourceWB As Object
Public g_SourceWS As Object
Private g_colIndexDict As Object
Private g_patternDict As Object
Private g_RowIndex As Object

' Log buffer for improved performance
Private g_LogBuffer As Object
Private g_LogEntryCount As Long

' Array pool for reuse - using bucketed approach for O(1) access
Private g_ArrayPool As Object
' Track pool statistics
Private g_ArrayPoolStats As Object

' ==============================================================================
' CORE UTILITY FUNCTIONS
' ==============================================================================
' ======== HandleError ========
Public Sub HandleError(moduleName As String, functionName As String, Optional errorMessage As String = "")
    ' FASTEST: Single decision with no function calls until needed
    If LenB(errorMessage) = 0 Then
        Log functionName & " ERROR: " & Err.Description & " (Err# " & Err.Number & ")", "ERROR"
    Else
        Log errorMessage, "ERROR"
    End If
End Sub

' ======== Log & FlushLogBuffer - OPTIMIZED COMBINED ========
Public Sub Log(msg As String, Optional sType As String = "INFO")
    On Error GoTo CriticalError
    
    ' Fast path: Log level filtering with static initialization
    Static levelMap As Object, minLevel As Long, initialized As Boolean
    Const LOG_DATE_FORMAT As String = "yyyy-mm-dd hh:mm:ss"
    
    If Not initialized Then
        Set levelMap = CreateObject("Scripting.Dictionary")
        levelMap.CompareMode = vbTextCompare
        levelMap("DEBUG") = 1: levelMap("INFO") = 2: levelMap("SUCCESS") = 2
        levelMap("WARNING") = 3: levelMap("ERROR") = 4: levelMap("FATAL") = 5
        
        Select Case UCase$(MIN_LOG_LEVEL)
            Case "DEBUG": minLevel = 1
            Case "INFO", "SUCCESS": minLevel = 2
            Case "WARNING": minLevel = 3
            Case "ERROR": minLevel = 4
            Case "FATAL": minLevel = 5
            Case Else: minLevel = 2
        End Select
        
        initialized = True
    End If
    
    ' Level check and exit if below minimum
    Dim msgLevel As Long: msgLevel = levelMap(UCase$(sType))
    If msgLevel = 0 Then msgLevel = 2
    If msgLevel < minLevel Then Exit Sub
    
    ' Create formatted message
    Dim logMsg As String: logMsg = Format(Now, LOG_DATE_FORMAT) & " [" & sType & "] " & msg
    
    ' Debug output
    If DEBUG_PRINT Then Debug.Print logMsg
    
    ' Buffer or direct write
    If Not g_LogBuffer Is Nothing And Not g_LogStream Is Nothing Then
        g_LogBuffer(g_LogEntryCount) = logMsg
        g_LogEntryCount = g_LogEntryCount + 1
        
        ' Flush conditions: critical errors or buffer full
        If sType = "ERROR" Or sType = "FATAL" Or g_LogEntryCount >= LOG_BUFFER_SIZE Then
            LogFlushBuffer
        End If
    ElseIf Not g_LogStream Is Nothing Then
        ' Direct write fallback
        g_LogStream.WriteLine logMsg
        If sType = "ERROR" Or sType = "FATAL" Then g_LogStream.Flush
    End If
    
    Exit Sub

CriticalError:
    If DEBUG_PRINT Then
        Debug.Print Format(Now, LOG_DATE_FORMAT) & " [LOG ERROR] " & _
                   Err.Description & " (Err# " & Err.Number & ")"
    End If
End Sub

' ======== LogFlushBuffer - FAST FLUSH ========
Private Sub LogFlushBuffer()
    On Error GoTo FlushError
    
    ' Fast validation
    If g_LogBuffer Is Nothing Then Exit Sub
    If g_LogStream Is Nothing Then Exit Sub
    If g_LogEntryCount = 0 Then Exit Sub
    
    ' Batch write all buffered entries
    Dim i As Long
    For i = 0 To g_LogEntryCount - 1
        g_LogStream.WriteLine g_LogBuffer(i)
    Next i
    
    ' Clear buffer and force disk write
    g_LogBuffer.RemoveAll
    g_LogEntryCount = 0
    g_LogStream.Flush
    
    Exit Sub

FlushError:
    ' Minimal error handling to prevent crash
    If DEBUG_PRINT Then
        Debug.Print Format(Now, "hh:mm:ss") & " [FLUSH ERROR] " & Err.Description
    End If
    ' Reset buffer to prevent repeated errors
    If Not g_LogBuffer Is Nothing Then
        g_LogBuffer.RemoveAll
        g_LogEntryCount = 0
    End If
End Sub

' ======== SafeArray2D - OPTIMIZED ========
Public Function SafeArray2D(inputValue As Variant) As Variant
    On Error GoTo ErrorHandler
    
    ' Handle Range objects
    If TypeName(inputValue) = "Range" Then
        SafeArray2D = inputValue.Value2
    Else
        SafeArray2D = inputValue
    End If
    
    ' Ensure it's a 2D array
    If Not IsArray(SafeArray2D) Then
        Dim singleValue(1 To 1, 1 To 1) As Variant
        singleValue(1, 1) = SafeArray2D
        SafeArray2D = singleValue
        Exit Function
    End If
    
    ' Check array dimensions (fails fast if not array)
    Dim dims As Integer
    dims = ArrayDimensions(SafeArray2D)
    
    If dims = 0 Then
        ' Not an array (shouldn't happen due to IsArray check, but defensive)
        Dim singleValue2(1 To 1, 1 To 1) As Variant
        singleValue2(1, 1) = SafeArray2D
        SafeArray2D = singleValue2
    ElseIf dims = 1 Then
        ' Convert 1D to 2D
        Dim arrayLowerBound As Long, arrayUpperBound As Long
        arrayLowerBound = LBound(SafeArray2D)
        arrayUpperBound = UBound(SafeArray2D)
        
        Dim result() As Variant
        ReDim result(1 To 1, 1 To arrayUpperBound - arrayLowerBound + 1)
        Dim i As Long, j As Long
        
        For i = arrayLowerBound To arrayUpperBound
            result(1, i - arrayLowerBound + 1) = SafeArray2D(i)
        Next i
        
        SafeArray2D = result
    ' If dims = 2, already correct, do nothing
    End If
    
    Exit Function
    
ErrorHandler:
    ' Return safe empty array on error
    Dim errArray(1 To 1, 1 To 1) As Variant
    errArray(1, 1) = CVErr(xlErrNA)
    SafeArray2D = errArray
    Log "SafeArray2D ERROR: " & Err.Description & " (Err# " & Err.Number & ")", "ERROR"
End Function

' ======== ArrayDimensions - OPTIMIZED WITH ERROR LOGGING ========
' ======== ArrayDimensions - OPTIMIZED & FIXED ========
Private Function ArrayDimensions(arr As Variant) As Integer
    ' ULTRA-FAST array dimension detection with minimal error handling
    On Error GoTo ErrorHandler
    
    ' Fast exit: Not an array
    If Not IsArray(arr) Then Exit Function
    
    Dim i As Integer: i = 1
    Dim temp As Long ' For UBound test
    
    On Error Resume Next ' Use inline error handling for speed
    
    Do While True
        ' Test dimension by attempting to get UBound
        temp = UBound(arr, i)
        If Err.Number = 9 Then ' Subscript out of range (expected - no more dimensions)
            Err.Clear
            Exit Do
        ElseIf Err.Number <> 0 Then ' Unexpected error
            ' Log and exit immediately
            Log "ArrayDimensions UNEXPECTED ERROR #" & Err.Number & ": " & Err.Description, "ERROR"
            ArrayDimensions = 0
            Exit Function
        End If
        i = i + 1
    Loop
    
    On Error GoTo ErrorHandler ' Restore normal error handling
    ArrayDimensions = i - 1
    Exit Function
    
ErrorHandler:
    ' Only log critical errors that bypassed inline handling
    Log "ArrayDimensions CRITICAL ERROR: " & Err.Description & " (Err# " & Err.Number & ")", "ERROR"
    ArrayDimensions = 0
End Function

' ======== GetPooledArray - OPTIMIZED & FIXED ========
Public Function GetPooledArray(Optional rows As Long = 1, Optional cols As Long = 1) As Variant
    On Error GoTo ErrorHandler
    
    ' Validate dimensions (minimum 1)
    If rows < 1 Then rows = 1
    If cols < 1 Then cols = 1
    
    ' Initialize pool if needed
    If g_ArrayPool Is Nothing Then InitializeArrayPool
    
    ' Fast path: Try to get from pool
    Dim bucketKey As String: bucketKey = rows & "x" & cols
    
    If g_ArrayPool.Exists(bucketKey) Then
        Dim bucket As Object: Set bucket = g_ArrayPool(bucketKey)
        If bucket.Count > 0 Then
            ' Get last item (fastest removal)
            Dim pooledArray As Variant
            pooledArray = bucket(bucket.Count - 1)
            bucket.Remove bucket.Count - 1
            
            ' Clear the array but keep dimensions
            Erase pooledArray
            GetPooledArray = pooledArray
            Exit Function
        End If
    End If
    
    ' Create new array if pool is empty
    Dim newArray() As Variant
    ReDim newArray(1 To rows, 1 To cols)
    GetPooledArray = newArray
    
    Exit Function
    
ErrorHandler:
    Log "GetPooledArray ERROR: " & Err.Description & " (Err# " & Err.Number & ")", "ERROR"
    
    ' Fail-safe: Return basic 1x1 array
    Dim failArray(1 To 1, 1 To 1) As Variant
    GetPooledArray = failArray
End Function

' ======== ReturnArrayToPool - OPTIMIZED ========
Public Sub ReturnArrayToPool(arr As Variant)
    On Error GoTo ErrorHandler
    
    ' Validate input and pool
    If Not IsArray(arr) Or IsError(arr) Then Exit Sub
    If g_ArrayPool Is Nothing Then Exit Sub
    
    ' Get array dimensions
    Dim rows As Long, cols As Long
    On Error Resume Next
    rows = UBound(arr, 1) - LBound(arr, 1) + 1
    cols = UBound(arr, 2) - LBound(arr, 2) + 1
    If Err.Number <> 0 Then Exit Sub
    On Error GoTo ErrorHandler
    
    ' Check pool size limit
    If g_ArrayPool.Count >= ARRAY_POOL_MAX_SIZE Then Exit Sub
    
    ' Add to pool
    Dim bucketKey As String: bucketKey = rows & "x" & cols
    
    If Not g_ArrayPool.Exists(bucketKey) Then
        Dim newBucket As Object
        Set newBucket = CreateObject("System.Collections.ArrayList")
        g_ArrayPool.Add bucketKey, newBucket
    End If
    
    ' Add array if bucket not full
    Dim bucket As Object: Set bucket = g_ArrayPool(bucketKey)
    If bucket.Count < ARRAY_POOL_TARGET_SIZE Then
        bucket.Add arr
    End If
    
    Exit Sub
    
ErrorHandler:
    Log "ReturnArrayToPool ERROR: " & Err.Description & " (Err# " & Err.Number & ")", "ERROR"
End Sub

' ======== InitializeArrayPool - OPTIMIZED ========
Private Sub InitializeArrayPool()
    On Error GoTo ErrorHandler
    
    If g_ArrayPool Is Nothing Then
        Set g_ArrayPool = CreateObject("Scripting.Dictionary")
        Log "Array pool initialized", "DEBUG"
    End If
    
    Exit Sub
    
ErrorHandler:
    Log "InitializeArrayPool ERROR: " & Err.Description & " (Err# " & Err.Number & ")", "ERROR"
    ' Don't re-raise - let caller handle missing pool gracefully
End Sub

' ============ MAIN EXECUTION - OPTIMIZED ============
Public Sub Calculation()
    Dim tTotal As Double: tTotal = Timer
    Dim bSuccess As Boolean: bSuccess = False
    
    On Error GoTo MainErr
    
    ' Initialize all global objects first (this is now the only required function to run)
    InitializeGlobalObjects
    
    ToggleOptimization True
    
    Log "========== PROCESS STARTED =========="
    
    If LoadSourceData() Then
        Dim lr As Long: lr = GetLastRow(g_SourceWS)
        
        If lr > 1 Then
            ProcessBatches
            bSuccess = True
        Else
            Log "No data rows to process (only header or empty)", "WARNING"
        End If
    End If
    
CleanExit:
    Dim duration As Double: duration = Timer - tTotal
    CleanUpResources
    ToggleOptimization False
    
    If bSuccess Then
        Log "========== COMPLETED | Total: " & Format(duration, "0.000") & "s ==========", "SUCCESS"
    Else
        Log "========== FAILED | Total: " & Format(duration, "0.000") & "s ==========", "ERROR"
    End If
    
    MsgBox IIf(bSuccess, "Process completed successfully!", "Process failed - check log"), _
           IIf(bSuccess, vbInformation, vbCritical), "Process Complete"
    Exit Sub

MainErr:
    Log "FATAL ERROR: " & Err.Description & " (Err# " & Err.Number & ")", "FATAL"
    bSuccess = False
    Resume CleanExit
End Sub

' ==============================================================================
' INITIALIZATION FUNCTIONS
' ==============================================================================

' ======== InitializeGlobalObjects - OPTIMIZED ========
Private Sub InitializeGlobalObjects()
    On Error GoTo ErrorHandler
    
    ' Initialize essential objects only
    If g_FSO Is Nothing Then Set g_FSO = CreateObject("Scripting.FileSystemObject")
    If g_HeaderMap Is Nothing Then Set g_HeaderMap = CreateObject("Scripting.Dictionary")
    If g_LookupData Is Nothing Then Set g_LookupData = CreateObject("Scripting.Dictionary")
    If g_RowIndex Is Nothing Then Set g_RowIndex = CreateObject("Scripting.Dictionary")
    If g_patternDict Is Nothing Then Set g_patternDict = CreateObject("Scripting.Dictionary")
    If g_LogBuffer Is Nothing Then Set g_LogBuffer = CreateObject("System.Collections.ArrayList")
    If g_ArrayPool Is Nothing Then Set g_ArrayPool = CreateObject("Scripting.Dictionary")
    
    ' Set compare modes once (performance optimization)
    g_HeaderMap.CompareMode = vbTextCompare
    g_LookupData.CompareMode = vbTextCompare
    g_RowIndex.CompareMode = vbTextCompare
    g_patternDict.CompareMode = vbTextCompare
    
    ' Initialize process ID if needed
    If Len(g_ProcessID) = 0 Then g_ProcessID = Format(Now, "YYYYMMDD_HHMMSS")
    
    ' Initialize log buffer (use Dictionary for performance)
    If g_LogBuffer Is Nothing Then
        Set g_LogBuffer = CreateObject("Scripting.Dictionary")
        g_LogEntryCount = 0
    End If
    
    ' Create required folders
    CreateFolderIfNotExists OUT_PATH, "Output"
    CreateFolderIfNotExists LOG_PATH, "Log"
    
    ' Create log file (only if not already created)
    If g_LogStream Is Nothing Then
        Dim logFile As String
        logFile = LOG_PATH & "Run_" & g_ProcessID & ".txt"
        
        Set g_LogStream = g_FSO.CreateTextFile(logFile, True)
        If DEBUG_PRINT Then Log "Log file created: " & logFile, "DEBUG"
    End If
    
    Exit Sub
    
ErrorHandler:
    Log "InitializeGlobalObjects ERROR: " & Err.Description & " (Err# " & Err.Number & ")", "ERROR"
    ' Re-raise error to ensure caller knows initialization failed
    Err.Raise Err.Number, "InitializeGlobalObjects", Err.Description
End Sub

' ======== SetSourceWorkbook - OPTIMIZED ========
Private Sub SetSourceWorkbook(wb As Workbook, ws As Worksheet)
    On Error GoTo ErrorHandler
    
    ' Fast path: validate inputs first
    If wb Is Nothing Or ws Is Nothing Then
        Log "SetSourceWorkbook ERROR: Invalid workbook or worksheet", "ERROR"
        Exit Sub
    End If
    
    ' Initialize only if needed (performance optimization)
    If g_HeaderMap Is Nothing Then InitializeGlobalObjects
    
    ' Set source references
    Set g_SourceWB = wb
    Set g_SourceWS = ws
    
    ' Log successful assignment in debug mode only
    If DEBUG_PRINT Then Log "Source workbook set: " & wb.Name, "DEBUG"
    
    Exit Sub
    
ErrorHandler:
    Log "SetSourceWorkbook ERROR: " & Err.Description & " (Err# " & Err.Number & ")", "ERROR"
End Sub

' ==============================================================================
' INITIALIZATION FUNCTIONS
' ==============================================================================

' ======== InitializeGlobals - OPTIMIZED ========
Private Sub InitializeGlobals()
    ' This function is kept for backward compatibility but now just calls the same initialization
    ' so users have the option to run it if needed, though it's not necessary anymore
    InitializeGlobalObjects
    Log "InitializeGlobals called (backward compatibility)", "INFO"
End Sub

' ======== CreateFolderIfNotExists - OPTIMIZED ========
Private Sub CreateFolderIfNotExists(folderPath As String, folderDescription As String)
    On Error GoTo ErrorHandler
    
    ' Fast path: validate inputs
    If Len(folderPath) = 0 Then Exit Sub
    
    ' Check if folder exists first (performance optimization)
    If Not g_FSO.FolderExists(folderPath) Then
        g_FSO.CreateFolder folderPath
        If DEBUG_PRINT Then Log folderDescription & " folder created: " & folderPath, "DEBUG"
    End If
    
    Exit Sub
    
ErrorHandler:
    Log "CreateFolderIfNotExists ERROR (" & folderDescription & "): " & _
        Err.Description & " (Err# " & Err.Number & ")", "ERROR"
End Sub

' ======== CleanUpResources - OPTIMIZED ========
Private Sub CleanUpResources()
    On Error GoTo ErrorHandler
    
    ' Flush logs first (critical for debugging)
    LogFlushBuffer
    
    ' Close and clear log stream
    If Not g_LogStream Is Nothing Then
        g_LogStream.Close
        Set g_LogStream = Nothing
    End If
    
    ' Close source workbook (if still open)
    If Not g_SourceWB Is Nothing Then
        g_SourceWB.Close False
        Set g_SourceWB = Nothing
    End If
    
    ' Clear remaining global objects
    Set g_FSO = Nothing
    Set g_HeaderMap = Nothing
    Set g_LookupData = Nothing
    Set g_patternDict = Nothing
    Set g_SourceWS = Nothing
    Set g_LogBuffer = Nothing
    
    ' Clean array pool efficiently
    If Not g_ArrayPool Is Nothing Then
        Dim key As Variant, bucket As Object
        
        For Each key In g_ArrayPool.Keys
            Set bucket = g_ArrayPool(key)
            bucket.Clear
            Set bucket = Nothing
        Next key
        
        g_ArrayPool.RemoveAll
        Set g_ArrayPool = Nothing
    End If
    
    ' Reset counters
    g_LogEntryCount = 0
    
    If DEBUG_PRINT Then Log "Resources cleaned up", "DEBUG"
    Exit Sub
    
ErrorHandler:
    Log "CleanUpResources ERROR: " & Err.Description & " (Err# " & Err.Number & ")", "ERROR"
    ' Continue cleanup despite errors
    Resume Next
End Sub

' ======== ToggleOptimization - OPTIMIZED ========
Private Sub ToggleOptimization(bOn As Boolean)
    On Error GoTo ErrorHandler
    
    With Application
        .ScreenUpdating = Not bOn
        .Calculation = IIf(bOn, xlCalculationManual, xlCalculationAutomatic)
        .EnableEvents = Not bOn
        .DisplayAlerts = Not bOn
        If bOn Then .StatusBar = False
    End With
    
    Exit Sub
    
ErrorHandler:
    Log "ToggleOptimization ERROR: " & Err.Description & " (Err# " & Err.Number & ")", "ERROR"
End Sub

' Calculate optimal chunk size based on system resources and data characteristics
' ======== CalculateOptimalChunkSize ========
Private Function CalculateOptimalChunkSize(colCount As Long, rowCount As Long) As Long
    ' FASTEST: Direct calculation with minimal branching
    
    ' Fast validation
    If rowCount <= 1 Then
        CalculateOptimalChunkSize = 1
        Exit Function
    End If
    
    ' Calculate base size based on column count
    Dim baseSize As Long
    
    ' Direct condition check (fastest possible)
    Select Case colCount
        Case 1 To 30:   baseSize = 100000  ' 100k rows for =30 cols
        Case 31 To 50:  baseSize = 50000   ' 50k rows for 31-50 cols
        Case Else:      baseSize = 25000   ' 25k rows for >50 cols
    End Select
    
    ' 64-bit systems can handle more memory
    #If Win64 Then
        baseSize = baseSize * 3 \ 2  ' Multiply by 1.5 using integer math (faster)
    #End If
    
    ' Limit to actual row count and ensure minimum
    CalculateOptimalChunkSize = Application.Min(baseSize, rowCount)
    If CalculateOptimalChunkSize < 1000 Then CalculateOptimalChunkSize = 1000
End Function


' ======== BrowseForFile ========
Private Function BrowseForFile() As String
    Static fd As FileDialog
    
    On Error GoTo ErrorHandler
    
    ' Create dialog once (performance optimization)
    If fd Is Nothing Then
        Set fd = Application.FileDialog(msoFileDialogFilePicker)
        With fd
            .Title = "Select Input File"
            .Filters.Clear
            .Filters.Add "Excel Files", "*.xlsx;*.xls;*.xlsm"
            .Filters.Add "CSV Files", "*.csv"
            .InitialFileName = ThisWorkbook.Path
        End With
    End If
    
    ' Show dialog and get result
    If fd.Show = -1 Then
        BrowseForFile = fd.SelectedItems(1)
    Else
        BrowseForFile = "" ' User cancelled
    End If
    
    Exit Function
    
ErrorHandler:
    Log "BrowseForFile ERROR: " & Err.Description & " (Err# " & Err.Number & ")", "ERROR"
    
    ' Reset dialog on error (force recreation next time)
    Set fd = Nothing
    BrowseForFile = ""
End Function

' ==============================================================================
' DATA LOADING FUNCTIONS
' ==============================================================================

' ======= LOAD SOURCE DATA ========
Private Function LoadSourceData() As Boolean
    On Error GoTo ErrorHandler
    
    LoadSourceData = False
    
    ' 1. Get file via dialog (static for performance)
    Static fd As FileDialog
    Dim fPath As String
    
    If fd Is Nothing Then
        Set fd = Application.FileDialog(msoFileDialogFilePicker)
        With fd
            .Title = "Select Input File"
            .Filters.Clear
            .Filters.Add "Excel Files", "*.xlsx;*.xls;*.xlsm"
            .Filters.Add "CSV Files", "*.csv"
            .InitialFileName = ThisWorkbook.Path
        End With
    End If
    
    If fd.Show = -1 Then
        fPath = fd.SelectedItems(1)
    Else
        Log "No file selected", "INFO"
        Exit Function
    End If
    
    ' 2. Fast validation
    If Len(fPath) = 0 Then Exit Function
    
    ' Fast file existence check (no FSO dependency)
    If Dir(fPath) = "" Then
        Log "File not found: " & fPath, "ERROR"
        Exit Function
    End If
    
    ' 3. Fast file type detection and opening
    Dim fileExt As String
    If InStrRev(fPath, ".") > 0 Then
        fileExt = LCase$(Mid$(fPath, InStrRev(fPath, "."))) ' Faster than Right$
    Else
        Log "Invalid file: no extension", "ERROR"
        Exit Function
    End If
    
    Log "Loading: " & fPath & " (" & UCase$(fileExt) & ")"
    
    ' Open workbook with appropriate settings
    On Error Resume Next
    Select Case fileExt
        Case ".csv"
            Set g_SourceWB = Workbooks.Open(Filename:=fPath, ReadOnly:=True, Local:=True)
        Case ".xlsx", ".xlsm", ".xls"
            Set g_SourceWB = Workbooks.Open(Filename:=fPath, ReadOnly:=True, UpdateLinks:=0)
        Case Else
            Log "Unsupported file type: " & fileExt, "ERROR"
            Exit Function
    End Select
    
    If Err.Number <> 0 Then
        Log "Failed to open " & UCase$(fileExt) & ": " & Err.Description, "ERROR"
        Set g_SourceWB = Nothing
        Exit Function
    End If
    On Error GoTo ErrorHandler
    
    ' 4. Get first worksheet (fast validation)
    If g_SourceWB Is Nothing Or g_SourceWB.Sheets.Count = 0 Then
        Log "Invalid workbook or no sheets", "ERROR"
        GoTo CleanUp
    End If
    
    Set g_SourceWS = g_SourceWB.Sheets(1)
    If TypeName(g_SourceWS) <> "Worksheet" Then
        Log "First sheet is not a worksheet", "ERROR"
        GoTo CleanUp
    End If
    
    ' 5. Quick dimension check (fast row count)
    Dim lr As Long: lr = GetLastRow(g_SourceWS)
    If lr < 1 Then
        Log "No data rows found", "ERROR"
        GoTo CleanUp
    End If
    
    ' 6. Log success and build header map
    Log "Loaded: " & lr & " rows × " & GetLastColumn(g_SourceWS) & " columns", "SUCCESS"
    RefreshHeaderMap g_SourceWS
    
    LoadSourceData = True
    Exit Function
    
CleanUp:
    ' Fast cleanup on failure
    If Not g_SourceWB Is Nothing Then
        g_SourceWB.Close False
        Set g_SourceWB = Nothing
    End If
    Set g_SourceWS = Nothing
    Exit Function
    
ErrorHandler:
    Log "LoadSourceData ERROR: " & Err.Description & " (Err# " & Err.Number & ")", "ERROR"
    LoadSourceData = False
End Function

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
    Dim lr1 As Long
    lr1 = ws.Cells(ws.rows.Count, 1).End(xlUp).Row
    
    ' Method 2: UsedRange (more reliable for CSV and modified files)
    Dim lr2 As Long
    lr2 = 0
    If Not usedRng Is Nothing Then
        lr2 = usedRng.Row + usedRng.rows.Count - 1
    End If
    
    ' Method 3: Direct UsedRange.Rows.Count
    Dim lr3 As Long
    lr3 = 0
    If Not usedRng Is Nothing Then
        lr3 = usedRng.rows.Count
    End If
    
    ' Method 4: Special handling for CSV - check if we have data beyond header
    Dim lr4 As Long
    lr4 = 0
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
            GetLastRow = ws.Cells(ws.rows.Count, 1).End(xlUp).Row
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
    For i = 1 To Application.Min(10, ws.rows.Count) ' Check first 10 rows or total rows if less
        candidateCol = ws.Cells(i, ws.Columns.Count).End(xlToLeft).Column
        If candidateCol > lastCol Then lastCol = candidateCol
    Next i
    
    ' Also check from bottom up to catch any data in far rows
    Dim lastRow As Long
    lastRow = ws.Cells(ws.rows.Count, 1).End(xlUp).Row
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
    
    Log "========== BATCH PROCESSING =========="
    
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
    HandleError "MainModule", "ProcessBatches"
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
    Dim tIndex As Double
    tIndex = Timer
    Application.ScreenUpdating = False
    BuildRowIndex g_SourceWS
    Application.ScreenUpdating = True
    Log "Row index built in " & Format(Timer - tIndex, "0.00") & "s"
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    HandleError "MainModule", "BuildMI_NORowIndex"
End Sub

' Calculate batch size based on total IDs
Private Function CalculateBatchSize(totalIDs As Long) As Long
    On Error GoTo ErrorHandler
    
    ' Define batch size based on total IDs
    If totalIDs <= 0 Then
        CalculateBatchSize = 1
    ElseIf totalIDs <= 100 Then
        CalculateBatchSize = totalIDs
    ElseIf totalIDs <= 1000 Then
        CalculateBatchSize = 100
    ElseIf totalIDs <= 10000 Then
        CalculateBatchSize = 500
    ElseIf totalIDs <= 100000 Then
        CalculateBatchSize = 1000
    Else
        CalculateBatchSize = 2000
    End If
    
    Exit Function
    
ErrorHandler:
    Log "Error in CalculateBatchSize: " & Err.Description & " - using default batch size", "WARNING"
    CalculateBatchSize = 1000 ' Default batch size
End Function

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
        ' Ceiling division for calculating total batches using native VBA
        CalculateTotalBatches = (totalIDs + batchSize - 1) \ batchSize
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
    Dim lastDoEventsTime As Single
    
    ' Initialize last DoEvents time
    lastDoEventsTime = Timer
    batchIdx = 0
    
    ' Process each batch
    For i = 1 To colUnique.Count Step batchSize
        Dim tBatch As Double
        tBatch = Timer
        batchIdx = batchIdx + 1
        
        ' Report progress every 10 batches or for the first and last batch
        If batchIdx Mod 10 = 1 Or batchIdx = totalBatches Or batchIdx = 1 Then
            Log "========== Batch " & batchIdx & "/" & totalBatches & " =========="
        End If
        
        ' Build batch dictionary
        Dim batchIDs As Object
        Set batchIDs = CreateObject("Scripting.Dictionary")
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
        
        ' Report details periodically to reduce log volume
        If batchIdx Mod 10 = 1 Or batchIdx = totalBatches Or batchIdx = 1 Then
            Log "  Contains " & batchIDs.Count & " MI_NOs (IDs " & i & " to " & endIdx & ")"
        End If
        
        ' Process batch
        If ProcessSingleBatch(batchIDs, batchIdx) Then
            ' Report completion periodically
            If batchIdx Mod 10 = 0 Or batchIdx = totalBatches Or batchIdx = 1 Then
                Log "  COMPLETED in " & Format(Timer - tBatch, "0.00") & "s"
            End If
        Else
            Log "  FAILED", "ERROR"
        End If
        
NextBatch:
        Set batchIDs = Nothing
        
        ' More intelligent DoEvents handling - based on time threshold or frequency
        If batchIdx Mod DOEVENTS_FREQUENCY = 0 Or (Timer - lastDoEventsTime) > DOEVENTS_TIME_THRESHOLD Then
            DoEvents
            lastDoEventsTime = Timer
        End If
        
        ' Update status bar with progress
        Application.StatusBar = "Processing batch " & batchIdx & " of " & totalBatches & " (" & Format(batchIdx / totalBatches * 100, "0.0") & "%)"
    Next i
    
    Application.StatusBar = False ' Reset status bar
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
    
    Dim filteredRows As Long
    filteredRows = GetLastRow(wsTemp)
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
    Set wbTemp = Workbooks.Add(-4167) ' xlWBATTemplate.xlWBATWorksheet
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
    
    ' Initialize global objects if not already done
    InitializeGlobalObjects
    
    FilterBatchData = False
    
    FilterDataToSheet g_SourceWS, wsTemp, batchIDs
    FilterBatchData = True
    
    Exit Function
    
FilterErr:
    Log "FILTER ERROR: " & Err.Description & " (Err# " & Err.Number & ")", "ERROR"
End Function


' Save batch output to file
' ======== SaveBatchOutput - ULTRA OPTIMIZED ========
Private Function SaveBatchOutput(wbTemp As Workbook, bIdx As Long) As Boolean
    Dim startTime As Double: startTime = Timer
    Dim fName As String, baseName As String
    Dim dotPos As Long, slashPos As Long
    
    On Error GoTo ErrorHandler
    
    ' Fast validation
    If wbTemp Is Nothing Then
        Log "SaveBatchOutput ERROR: Temporary workbook is Nothing", "ERROR"
        SaveBatchOutput = False
        Exit Function
    End If
    
    SaveBatchOutput = False
    
    ' Ensure output folder exists - single operation check
    If Len(Dir(OUT_PATH, vbDirectory)) = 0 Then
        MkDir OUT_PATH
        Log "Created output folder: " & OUT_PATH, "DEBUG"
    End If
    
    ' Extract base name without extension in one operation
    baseName = g_SourceWB.Name
    dotPos = InStrRev(baseName, ".")
    If dotPos > 0 Then baseName = Left$(baseName, dotPos - 1)
    
    ' Construct full filename in single operation
    fName = OUT_PATH & baseName & "_Batch_" & Format(bIdx, "000") & "_" & g_ProcessID & ".csv"
    
    ' Save with minimal operations
    Application.DisplayAlerts = False
    wbTemp.SaveAs fName, xlCSV
    Application.DisplayAlerts = True
    
    ' Extract just filename for logging
    slashPos = InStrRev(fName, "\")
    If slashPos = 0 Then slashPos = InStrRev(fName, "/")
    
    ' Log success with performance metrics
    Dim fileNameOnly As String
    If slashPos > 0 Then
        fileNameOnly = Mid$(fName, slashPos + 1)
    Else
        fileNameOnly = fName
    End If
    
    Log "Saved: " & fileNameOnly & " in " & Format(Timer - startTime, "0.000") & "s", "SUCCESS"
    
    SaveBatchOutput = True
    Exit Function
    
ErrorHandler:
    ' Ensure DisplayAlerts is restored even on error
    Application.DisplayAlerts = True
    
    ' Build error message with context
    Dim errMsg As String
    errMsg = "SAVE ERROR in batch " & bIdx & ": " & Err.Description
    
    ' Include attempted filename in error log (truncate if too long)
    If Len(fName) > 0 Then
        If Len(fName) > 100 Then
            errMsg = errMsg & " [Path: ..." & Right$(fName, 100) & "]"
        Else
            errMsg = errMsg & " [Path: " & fName & "]"
        End If
    End If
    
    Log errMsg & " (Error " & Err.Number & ")", "ERROR"
    SaveBatchOutput = False
End Function

' =============HELPER SYSTEM FUNCTIONS===============================================
' ======== RunAllHelpers ========
Private Sub RunAllHelpers(ws As Worksheet)
    Dim tStart As Double: tStart = Timer
    Dim helpers() As String ' String array for speed
    Dim i As Long, executed As Long
    Dim bSuccess As Boolean: bSuccess = True
    
    On Error GoTo ErrorHandler
    
    ' Build array line by line (fast and easy to comment)
    ReDim helpers(0 To 32) ' Adjust size as needed
    
    ' EACH HELPER ON SEPARATE LINE - EASY TO COMMENT
    helpers(0) = "Helper_01_SubGroupMapping"
    helpers(1) = "Helper_02_Source"
    helpers(2) = "Helper_03_NB_and_RI"
    helpers(3) = "Helper_04_DefaultRate"
    helpers(4) = "Helper_05_LTVPropValue"
    helpers(5) = "Helper_06_AssumptionValuation"
    helpers(6) = "Helper_07_Policies_Date"
    helpers(7) = "Helper_08_ScenarioMerge"
    helpers(8) = "Helper_09_ReinsurerMapping"
    helpers(9) = "Helper_10_AddGMR_MTHLY"
    helpers(10) = "Helper_11_AddMaturity_Term"
    helpers(11) = "Helper_13_Projection_Expand"
    helpers(12) = "Helper_14_AddRemaining_Term"
    helpers(13) = "Helper_15_YrMth_Indicator"
    helpers(14) = "Helper_16_DefaultPattern"
    helpers(15) = "Helper_17_ClaimRate"
    helpers(16) = "Helper_18_CalculateOMCR"
    helpers(17) = "Helper_20_OMV_and_PctOMV_Combined"
    helpers(18) = "Helper_21_CPR_SMM_Combined"
    helpers(19) = "Helper_23_AcquisitionExpense"
    helpers(20) = "Helper_24_PolicyForceAndOPB_Combined"
    helpers(21) = "Helper_24A_Duration_Plus_and_Min"
    helpers(22) = "Helper_24B_Next_Policy_In_Force"
    helpers(23) = "Helper_24C_PreviousValues"
    helpers(24) = "Helper_25_FixedAssumedSeverity"
    helpers(25) = "Helper_26_InflationFactor"
    helpers(26) = "Helper_27_MaintenanceExpense"
    helpers(27) = "Helper_28_Commission_and_Commission_Recovery"
    helpers(28) = "Helper_29_RI_Policies_Count"
    helpers(29) = "Helper_30_RI_Premium"
    helpers(30) = "Helper_31_RiskInForce_and_DefaultClaimOutgo"
    helpers(31) = "Helper_32_RI_Collateral"
    helpers(32) = "Helper_33_RI_NPR"
    
    Log "Starting " & (UBound(helpers) + 1) & " helpers...", "INFO"
    
    ' Execute helpers
    For i = LBound(helpers) To UBound(helpers)
        If LenB(helpers(i)) = 0 Then GoTo ContinueLoop ' Skip if commented out
        
        executed = executed + 1
        Dim tHelper As Double: tHelper = Timer
        
        ' Fast inline execution
        On Error Resume Next
        Application.Run helpers(i), ws
        
        If Err.Number <> 0 Then
            Log helpers(i) & " FAILED: " & Err.Description & " [" & _
                Format(Timer - tHelper, "0.000") & "s]", "ERROR"
            Err.Clear
            bSuccess = False
        Else
            Log helpers(i) & " completed in " & Format(Timer - tHelper, "0.000") & "s", "SUCCESS"
            
            ' Header refresh logic
            Select Case Mid$(helpers(i), 9, 2)
                Case "11", "20", "23"
                    RefreshHeaderMap ws
                    Log "Headers refreshed after " & helpers(i), "INFO"
            End Select
        End If
        On Error GoTo ErrorHandler
        
        ' Progress update
        If executed Mod 10 = 0 Then
            Log "Progress: " & executed & " of " & (UBound(helpers) + 1) & " completed", "INFO"
        End If
        
ContinueLoop:
    Next i
    
    ' Final performance log
    Log (UBound(helpers) + 1) & " helpers completed in " & _
        Format(Timer - tStart, "0.000") & "s", IIf(bSuccess, "SUCCESS", "WARNING")
    
    Exit Sub
    
ErrorHandler:
    Log "RunAllHelpers ERROR at index " & i & ": " & Err.Description, "ERROR"
    bSuccess = False
    Resume ContinueLoop
End Sub
' ======== ExecuteHelperSafe - ENHANCED VERSION ========
Private Function ExecuteHelperSafe(ws As Worksheet, helperName As String) As Boolean
    Dim startTime As Double: startTime = Timer
    
    On Error Resume Next
    Application.Run helperName, ws
    
    If Err.Number <> 0 Then
        ' Log error with duration
        Log helperName & " FAILED: " & Err.Description & _
            " (Error " & Err.Number & ") [" & Format(Timer - startTime, "0.000") & "s]", "ERROR"
        Err.Clear
        ExecuteHelperSafe = False
    Else
        ' Log success with duration
        Log helperName & " completed in " & Format(Timer - startTime, "0.000") & "s", "SUCCESS"
        ExecuteHelperSafe = True
    End If
    
    On Error GoTo 0
End Function
' ======== IsInArray - Helper function to check if a value exists in an array ========
Private Function IsInArray(value As Long, arr As Variant) As Boolean
    Dim i As Long, arrayLowerBound As Long, arrayUpperBound As Long
    
    arrayLowerBound = LBound(arr)
    arrayUpperBound = UBound(arr)
    
    For i = arrayLowerBound To arrayUpperBound
        If arr(i) = value Then
            IsInArray = True
            Exit Function
        End If
    Next i
    ' Return False (default) - no need to explicitly set
End Function

' ======== ExecuteHelper - OPTIMIZED ========
Private Sub ExecuteHelper(ws As Worksheet, procName As String)
    Dim startTime As Double: startTime = Timer
    On Error GoTo ErrorHandler
    
    Application.Run procName, ws
    Log "      " & procName & ": " & Format(Timer - startTime, "0.000") & "s", "SUCCESS"
    
    Exit Sub
    
ErrorHandler:
    Log "      " & procName & " ERROR(" & Err.Number & "): " & Err.Description & _
        " [Duration: " & Format(Timer - startTime, "0.000") & "s]", "ERROR"
    Err.Raise Err.Number, "ExecuteHelper", Err.Description
End Sub

' ===========FILTERING ENGINE=========================
' ======== BuildRowIndex - OPTIMIZED ========
Private Sub BuildRowIndex(ws As Worksheet)
    On Error GoTo IndexErr
    
    ' Fast validation
    If ws Is Nothing Then
        Set g_RowIndex = CreateObject("Scripting.Dictionary")
        Log "BuildRowIndex: Worksheet is Nothing", "ERROR"
        Exit Sub
    End If
    
    Dim startTime As Double: startTime = Timer
    Dim lr As Long: lr = GetLastRow(ws)
    
    ' Exit early if no data
    If lr < 2 Then
        Log "No data to index", "INFO"
        Exit Sub
    End If
    
    ' Initialize dictionary with binary compare for speed
    Set g_RowIndex = CreateObject("Scripting.Dictionary")
    g_RowIndex.CompareMode = vbBinaryCompare
    
    ' Single read operation for column A
    Dim vMI As Variant
    vMI = SafeArray2D(ws.Range(ws.Cells(2, 1), ws.Cells(lr, 1)).Value2)
    
    Dim rowCount As Long: rowCount = UBound(vMI, 1)
    Dim i As Long, miNo As String, rowCol As Collection
    
    ' Fast processing loop
    For i = 1 To rowCount
        miNo = Trim$(NzString(vMI, i, 1))
        
        ' Skip empty MI_NOs
        If LenB(miNo) > 0 Then
            If Not g_RowIndex.Exists(miNo) Then
                Set rowCol = New Collection
                rowCol.Add i + 1 ' Adjust for header row
                g_RowIndex.Add miNo, rowCol
            Else
                ' Add to existing collection
                Set rowCol = g_RowIndex(miNo)
                rowCol.Add i + 1
            End If
        End If
    Next i
    
    ' Cleanup
    Erase vMI
    Set rowCol = Nothing
    
    Log "Indexed " & g_RowIndex.Count & " unique MI_NOs in " & _
        Format(Timer - startTime, "0.000") & "s", "SUCCESS"
    
    Exit Sub
    
IndexErr:
    ' Ensure dictionary exists even on error
    If g_RowIndex Is Nothing Then
        Set g_RowIndex = CreateObject("Scripting.Dictionary")
    End If
    
    Log "BuildRowIndex ERROR: " & Err.Description & _
        " (Err# " & Err.Number & ")", "ERROR"
End Sub

'=================FilterDataToSheet========================
Private Sub FilterDataToSheet(srcWS As Worksheet, destWS As Worksheet, validIDs As Object)
    On Error GoTo FilterErr
    
    ' Fast validation
    If srcWS Is Nothing Or destWS Is Nothing Then
        Log "FilterDataToSheet: Invalid worksheet reference", "ERROR"
        Exit Sub
    End If
    
    Dim startTime As Double: startTime = Timer
    Dim lr As Long: lr = GetLastRow(srcWS)
    
    ' Copy header once
    srcWS.rows(1).Copy destWS.rows(1)
    
    ' Exit early if no data
    If lr < 2 Then
        Log "No data to filter", "INFO"
        Exit Sub
    End If
    
    Dim lc As Long: lc = GetLastColumn(srcWS)
    Dim vSrc As Variant, vRes As Variant
    Dim srcRows As Long, srcCols As Long, outR As Long
    Dim i As Long, j As Long, c As Long
    Dim miNo As String, rowCol As Collection
    
    ' Single read operation
    vSrc = SafeArray2D(srcWS.Range("A1", srcWS.Cells(lr, lc)).Value2)
    srcRows = UBound(vSrc, 1)
    srcCols = UBound(vSrc, 2)
    
    Dim srcLB2 As Long: srcLB2 = LBound(vSrc, 2)
    Dim srcUB2 As Long: srcUB2 = UBound(vSrc, 2)
    
    ' Pre-count matching rows for exact allocation
    Dim matchCount As Long: matchCount = 0
    
    If Not g_RowIndex Is Nothing And g_RowIndex.Count > 0 Then
        ' Fast path with index
        Dim key As Variant
        For Each key In validIDs.Keys
            miNo = Trim$(CStr(key))
            If LenB(miNo) > 0 And g_RowIndex.Exists(miNo) Then
                Set rowCol = g_RowIndex(miNo)
                matchCount = matchCount + rowCol.Count
                Set rowCol = Nothing
            End If
        Next key
    Else
        ' Fallback: scan source
        For i = 2 To srcRows
            miNo = Trim$(NzString(vSrc, i, 1))
            If LenB(miNo) > 0 And validIDs.Exists(miNo) Then
                matchCount = matchCount + 1
            End If
        Next i
    End If
    
    ' Exit early if no matches
    If matchCount = 0 Then
        Log "No matching data found", "INFO"
        Erase vSrc
        Exit Sub
    End If
    
    ' Allocate exact size using pooled array
    vRes = GetPooledArray(matchCount, srcCols)
    outR = 0
    
    If Not g_RowIndex Is Nothing And g_RowIndex.Count > 0 Then
        ' Fast indexed copy
        For Each key In validIDs.Keys
            miNo = Trim$(CStr(key))
            If LenB(miNo) > 0 And g_RowIndex.Exists(miNo) Then
                Set rowCol = g_RowIndex(miNo)
                
                For j = 1 To rowCol.Count
                    i = rowCol(j)
                    If i >= 2 And i <= srcRows Then
                        outR = outR + 1
                        ' Bulk copy row data
                        For c = srcLB2 To srcUB2
                            vRes(outR, c - srcLB2 + 1) = vSrc(i, c)
                        Next c
                    End If
                Next j
                Set rowCol = Nothing
            End If
        Next key
    Else
        ' Fallback copy
        For i = 2 To srcRows
            miNo = Trim$(NzString(vSrc, i, 1))
            If LenB(miNo) > 0 And validIDs.Exists(miNo) Then
                outR = outR + 1
                ' Bulk copy row data
                For c = srcLB2 To srcUB2
                    vRes(outR, c - srcLB2 + 1) = vSrc(i, c)
                Next c
            End If
        Next i
    End If
    
    ' Single write operation
    If outR > 0 Then
        destWS.Range("A2").Resize(outR, srcCols).Value2 = vRes
    End If
    
    ' Cleanup
    ReturnArrayToPool vRes
    Erase vSrc
    
    Log "FilterDataToSheet: " & outR & " rows filtered in " & _
        Format(Timer - startTime, "0.000") & "s", "SUCCESS"
    
    Exit Sub
    
FilterErr:
    ' Cleanup on error
    On Error Resume Next
    If IsArray(vRes) Then ReturnArrayToPool vRes
    Erase vSrc
    Set rowCol = Nothing
    
    Log "FilterDataToSheet ERROR: " & Err.Description & _
        " (Err# " & Err.Number & ")", "ERROR"
End Sub

' ======== GetDistinctIDs ========
Private Function GetDistinctIDs(ws As Worksheet) As Collection
    On Error GoTo ErrorHandler
    
    ' Fast validation
    If ws Is Nothing Then Set GetDistinctIDs = New Collection: Exit Function
    
    Dim lastRow As Long: lastRow = GetLastRow(ws)
    If lastRow < 2 Then Set GetDistinctIDs = New Collection: Exit Function
    
    ' Read ID column directly
    Dim ids As Variant
    ids = ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, 1)).Value2
    ids = SafeArray2D(ids)
    
    ' Use dictionary for uniqueness (faster than collection with error handling)
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare
    
    Dim i As Long, idText As String
    For i = 1 To UBound(ids, 1)
        idText = CStr(ids(i, 1))
        If Len(idText) > 0 Then
            dict(Trim$(idText)) = Empty
        End If
    Next i
    
    ' Convert dictionary keys to collection (fast bulk operation)
    Set GetDistinctIDs = New Collection
    Dim key As Variant
    For Each key In dict.Keys
        GetDistinctIDs.Add key
    Next key
    
    Exit Function
    
ErrorHandler:
    Log "GetDistinctIDs ERROR: " & Err.Description & " (Err# " & Err.Number & ")", "ERROR"
    Set GetDistinctIDs = New Collection
End Function

' ======== AddColumn - COLUMN MANAGEMENT FUNCTIONS ========
Public Sub AddColumn(ws As Worksheet, colName As String)
    On Error GoTo ErrorHandler
    
    ' Fast validation
    If ws Is Nothing Then Exit Sub
    If Len(colName) = 0 Then Exit Sub
    If g_HeaderMap Is Nothing Then RefreshHeaderMap ws
    
    ' Normalize column name once
    Dim ucName As String: ucName = UCase$(Trim$(colName))
    
    ' Fast check: Already exists in map
    If g_HeaderMap.Exists(ucName) Then Exit Sub
    
    ' Find next empty column
    Dim lastCol As Long: lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    Dim targetCol As Long: targetCol = lastCol + 1
    
    ' Write header and update map
    ws.Cells(1, targetCol).value = colName
    g_HeaderMap(ucName) = targetCol
    
    If DEBUG_PRINT Then Log "Added column: " & colName & " at " & targetCol, "DEBUG"
    Exit Sub
    
ErrorHandler:
    Log "AddColumn ERROR: " & Err.Description & " (Err# " & Err.Number & ")", "ERROR"
End Sub

' ======== GetColumnIndex ========
Public Function GetColumnIndex(ws As Worksheet, colName As String) As Long
    On Error GoTo ErrorHandler
    
    ' Fast validation
    If ws Is Nothing Then Exit Function
    If Len(colName) = 0 Then Exit Function
    If g_HeaderMap Is Nothing Then Exit Function
    
    ' Fast lookup with normalization
    Dim ucName As String: ucName = UCase$(Trim$(colName))
    
    If g_HeaderMap.Exists(ucName) Then
        GetColumnIndex = g_HeaderMap(ucName)
    End If
    
    Exit Function
    
ErrorHandler:
    Log "GetColumnIndex ERROR: " & Err.Description & " (Err# " & Err.Number & ")", "ERROR"
    GetColumnIndex = 0
End Function

' ======== RefreshHeaderMap ========
Private Sub RefreshHeaderMap(ws As Worksheet)
    On Error GoTo ErrorHandler
    
    ' Fast validation
    If ws Is Nothing Then Exit Sub
    
    ' Clear existing map
    g_HeaderMap.RemoveAll
    
    ' Read header row directly (fastest method)
    Dim headers As Variant
    headers = ws.rows(1).Value2
    headers = SafeArray2D(headers)
    
    ' Fast iteration through columns
    Dim col As Long, lastCol As Long
    lastCol = UBound(headers, 2)
    
    For col = 1 To lastCol
        Dim headerVal As Variant: headerVal = headers(1, col)
        
        ' Skip errors and empties
        If Not IsError(headerVal) Then
            If Not IsEmpty(headerVal) Then
                ' Fast string conversion and normalization
                Dim headerText As String: headerText = CStr(headerVal)
                If Len(headerText) > 0 Then
                    g_HeaderMap(UCase$(Trim$(headerText))) = col
                End If
            End If
        End If
    Next col
    
    Exit Sub
    
ErrorHandler:
    Log "RefreshHeaderMap ERROR: " & Err.Description & " (Err# " & Err.Number & ")", "ERROR"
End Sub

' ======== ExecuteBatchHelpers ========
Private Sub ExecuteBatchHelpers(ws As Worksheet)
    On Error GoTo ErrorHandler
    
    Log "Starting batch helpers execution...", "INFO"
    
    ' Execute all helpers defined in the RunAllHelpers function
    RunAllHelpers ws
    
    Log "Batch helpers execution completed", "INFO"
    Exit Sub
    
ErrorHandler:
    Log "ExecuteBatchHelpers ERROR: " & Err.Description & " (Err# " & Err.Number & ")", "ERROR"
End Sub

' ==============================================================================
' TYPE-SAFE CONVERSION FUNCTIONS
' ==============================================================================
' ======== NzLong ========
Public Function NzLong(data As Variant, r As Long, c As Long) As Long
    On Error Resume Next
    
    ' Fast path: Direct array access
    Dim v As Variant: v = data(r, c)
    
    If Err.Number = 0 Then
        If Not IsError(v) Then
            If Not IsEmpty(v) And Not IsNull(v) Then
                If IsNumeric(v) Then
                    NzLong = CLng(v)
                End If
            End If
        End If
    End If
    
    On Error GoTo 0
End Function

' ======== NzDouble ========
Public Function NzDouble(data As Variant, r As Long, c As Long) As Double
    On Error Resume Next
    
    ' Fast path: Direct array access
    Dim v As Variant: v = data(r, c)
    
    If Err.Number = 0 Then
        If Not IsError(v) Then
            If Not IsEmpty(v) And Not IsNull(v) Then
                If IsNumeric(v) Then
                    NzDouble = CDbl(v)
                End If
            End If
        End If
    End If
    
    On Error GoTo 0
End Function

' ======== NzString ========
Public Function NzString(data As Variant, r As Long, c As Long) As String
    On Error Resume Next
    
    ' Fast path: Direct array access with minimal checks
    Dim v As Variant: v = data(r, c)
    
    If Err.Number = 0 Then
        If Not IsError(v) Then
            If Not IsEmpty(v) And Not IsNull(v) Then
                NzString = Trim$(CStr(v))
            End If
        End If
    End If
    
    On Error GoTo 0
End Function


' ======== ParseDate - HANDLES YYYYMMDD FORMAT ========
Public Function ParseDate(ByVal v As Variant, ByRef outDate As Date) As Boolean
    On Error GoTo ErrorHandler
    
    ' FAST: Empty/Null/Error check
    If IsEmpty(v) Or IsNull(v) Or IsError(v) Then Exit Function
    
    ' FAST: Already a date (quickest path)
    If VarType(v) = vbDate Then
        outDate = v
        ParseDate = True
        Exit Function
    End If
    
    ' MEDIUM: Numeric input (YYYYMMDD or Excel serial)
    If IsNumeric(v) Then
        Dim dblVal As Double: dblVal = CDbl(v)
        
        ' YYYYMMDD format (fast integer math)
        If dblVal >= 19000101 And dblVal <= 99991231 Then
            Dim y As Long: y = dblVal \ 10000           ' Fast integer division
            Dim m As Long: m = (dblVal Mod 10000) \ 100 ' Use Mod for speed
            Dim d As Long: d = dblVal Mod 100
            
            outDate = DateSerial(y, m, d)
            ' Verify conversion (defensive)
            If Year(outDate) = y And Month(outDate) = m And Day(outDate) = d Then
                ParseDate = True
            End If
            Exit Function
        End If
        
        ' Excel serial date
        If dblVal > 0 And dblVal < 2958466 Then
            outDate = CDate(dblVal)
            ParseDate = True
            Exit Function
        End If
    End If
    
    ' SLOW: Text date (CSV often loads as text) - only if nothing else worked
    If VarType(v) = vbString And Len(v) > 0 Then
        Dim dateStr As String: dateStr = CStr(v)
        
        ' Try direct conversion first (fastest for text)
        On Error Resume Next
        outDate = CDate(dateStr)
        If Err.Number = 0 Then
            ParseDate = True
            On Error GoTo ErrorHandler
            Exit Function
        End If
        On Error GoTo ErrorHandler
        
        ' Try common date formats (ordered by frequency)
        Dim formats As Variant
        formats = Array("yyyy-mm-dd", "mm/dd/yyyy", "dd/mm/yyyy", "yyyy/mm/dd")
        
        Dim formatLowerBound As Long, formatUpperBound As Long
        formatLowerBound = LBound(formats)
        formatUpperBound = UBound(formats)
        
        Dim i As Long
        For i = formatLowerBound To formatUpperBound
            On Error Resume Next
            outDate = CDate(Format(dateStr, formats(i)))
            If Err.Number = 0 Then
                ParseDate = True
                On Error GoTo ErrorHandler
                Exit Function
            End If
            On Error GoTo ErrorHandler
        Next i
    End If
    
    Exit Function
    
ErrorHandler:
    Log "ParseDate ERROR: " & Err.Description & " | Value: " & CStr(v), "ERROR"
    ParseDate = False
End Function

' ======== TryConvertVariant ========
' Helper function to safely convert variant to double, currency or long
Private Function TryConvertVariant(value As Variant, ByRef result As Variant, targetType As String) As Boolean
    ' FASTEST: Direct conversion with type-specific error handling
    
    Select Case targetType
        Case "Double"
            On Error Resume Next
            result = CDbl(value)
            TryConvertVariant = (Err.Number = 0)
            On Error GoTo 0
        Case "Long"
            On Error Resume Next
            result = CLng(value)
            TryConvertVariant = (Err.Number = 0)
            On Error GoTo 0
        Case "Currency"
            On Error Resume Next
            result = CCur(value)
            TryConvertVariant = (Err.Number = 0)
            On Error GoTo 0
        Case "String"
            On Error Resume Next
            result = CStr(value)
            TryConvertVariant = (Err.Number = 0)
            On Error GoTo 0
        Case Else
            TryConvertVariant = False
            Exit Function
    End Select
    
    ' Log only if conversion failed (fail fast with logging)
    If Not TryConvertVariant Then
        Log "TryConvertVariant ERROR: Cannot convert to " & targetType & " - " & _
            TypeName(value) & " = " & Left$(CStr(value), 100), "WARNING"
    End If
End Function
' ==============================================================================
' END OF MAIN MODULE
' ==============================================================================



' ==============================================================================
' HELPER FUNCTIONS
' ==============================================================================

