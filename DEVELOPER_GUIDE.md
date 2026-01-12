# RMP Excel VBA Processing Engine - Developer Guide

## Overview

This developer guide documents the core utility functions available in the RMP Excel VBA Processing Engine that can be shared and used throughout the model. These functions provide robust error handling, performance optimization, and consistent behavior across different parts of the application.

## Table of Contents

1. [Data Type Conversion Functions](#data-type-conversion-functions)
2. [Array Handling Functions](#array-handling-functions)
3. [Column Management Functions](#column-management-functions)
4. [Date Handling Functions](#date-handling-functions)
5. [Memory Management Functions](#memory-management-functions)
6. [Processing Functions](#processing-functions)
7. [Logging and Error Handling](#logging-and-error-handling)
8. [Dictionary and Collection Functions](#dictionary-and-collection-functions)

---

## Data Type Conversion Functions

### NzLong

Safely converts a value from a 2D array to Long with comprehensive error handling.

```vb
Public Function NzLong(data As Variant, row As Long, col As Long) As Long
```

**Parameters:**
- `data`: The 2D variant array containing the source data
- `row`: Row index to access in the array
- `col`: Column index to access in the array

**Returns:**
- Long value from the specified array element, or 0 if conversion fails

**Example:**
```vb
Dim myArray As Variant
myArray = Range("A1:C10").Value2
Dim myLongValue As Long
myLongValue = NzLong(myArray, 3, 2) ' Get value from row 3, column 2 as Long
```

### NzDouble

Safely converts a value from a 2D array to Double with comprehensive error handling.

```vb
Public Function NzDouble(data As Variant, r As Long, c As Long) As Double
```

**Parameters:**
- `data`: The 2D variant array containing the source data
- `r`: Row index to access in the array
- `c`: Column index to access in the array

**Returns:**
- Double value from the specified array element, or 0 if conversion fails

**Example:**
```vb
Dim myArray As Variant
myArray = Range("A1:C10").Value2
Dim myDoubleValue As Double
myDoubleValue = NzDouble(myArray, 3, 2) ' Get value from row 3, column 2 as Double
```

### NzString

Safely converts a value from a 2D array to String with comprehensive error handling.

```vb
Public Function NzString(data As Variant, r As Long, c As Long) As String
```

**Parameters:**
- `data`: The 2D variant array containing the source data
- `r`: Row index to access in the array
- `c`: Column index to access in the array

**Returns:**
- String value from the specified array element, or empty string if conversion fails

**Example:**
```vb
Dim myArray As Variant
myArray = Range("A1:C10").Value2
Dim myStringValue As String
myStringValue = NzString(myArray, 3, 2) ' Get value from row 3, column 2 as String
```

### TryConvertVariant

Attempts to convert a variant to a specified data type with error handling.

```vb
Private Function TryConvertVariant(value As Variant, ByRef result As Variant, targetType As String) As Boolean
```

**Parameters:**
- `value`: The variant value to convert
- `result`: Reference variable that receives the converted result
- `targetType`: String indicating the target type ("Double", "Long", or "Currency")

**Returns:**
- Boolean indicating whether conversion succeeded

**Example:**
```vb
Dim myResult As Variant
Dim success As Boolean
success = TryConvertVariant(Range("A1").Value, myResult, "Double")
If success Then
    ' Use myResult safely as Double
End If
```

---

## Array Handling Functions

### SafeArray2D

Ensures a variant is a valid 2D array, converting if necessary.

```vb
Private Function SafeArray2D(inputValue As Variant) As Variant
```

**Parameters:**
- `inputValue`: The input variant that might be a 1D array, 2D array, or scalar value

**Returns:**
- A 2D array version of the input

**Example:**
```vb
Dim myData As Variant
myData = Range("A1").Value ' Might be a scalar
Dim my2DArray As Variant
my2DArray = SafeArray2D(myData) ' Now guaranteed to be a 2D array
```

---

## Column Management Functions

### AddColumn

Adds a column to a worksheet with comprehensive error handling and recovery options.

```vb
Public Sub AddColumn(ws As Worksheet, colName As String)
```

**Parameters:**
- `ws`: Target worksheet
- `colName`: Name of the column header to add

**Example:**
```vb
AddColumn Sheets("Data"), "Commission_Rate"
```

### GetColumnIndex

Gets the column index for a named column in a worksheet.

```vb
Public Function GetColumnIndex(ws As Worksheet, colName As String) As Long
```

**Parameters:**
- `ws`: Target worksheet
- `colName`: Name of the column to find

**Returns:**
- Column index (1-based) or 0 if not found

**Example:**
```vb
Dim colIdx As Long
colIdx = GetColumnIndex(Sheets("Data"), "Amount")
If colIdx > 0 Then
    ' Column exists, use the index
End If
```

### RefreshHeaderMap

Rebuilds the internal header map dictionary for faster column lookups.

```vb
Private Sub RefreshHeaderMap(ws As Worksheet)
```

**Parameters:**
- `ws`: Target worksheet

**Example:**
```vb
' Call before performing operations requiring column lookups
RefreshHeaderMap Sheets("Data")
```

---

## Date Handling Functions

### ParseDate

Parses dates from various formats, including YYYYMMDD numeric format.

```vb
Public Function ParseDate(ByVal v As Variant, ByRef outDate As Date) As Boolean
```

**Parameters:**
- `v`: Value to parse as date
- `outDate`: Reference variable that receives the parsed date

**Returns:**
- Boolean indicating whether parsing succeeded

**Example:**
```vb
Dim myDate As Date
Dim success As Boolean
success = ParseDate(20231225, myDate) ' Parse numeric YYYYMMDD format
If success Then
    ' Use myDate safely
End If
```

---

## Memory Management Functions

### GetAvailableMemoryMB

Estimates available system memory for optimizing chunking operations.

```vb
Private Function GetAvailableMemoryMB() As Long
```

**Returns:**
- Estimated available memory in MB, or 0 if unknown

**Example:**
```vb
Dim availableMem As Long
availableMem = GetAvailableMemoryMB()
If availableMem < 100 Then
    ' Use smaller chunk sizes for processing
End If
```

### GetPooledArray

Gets an array from the memory pool to reduce memory fragmentation.

```vb
' Referenced in code but implementation not visible
' Usage pattern based on code references:
Dim vRes As Variant
vRes = GetPooledArray(rows, columns, lowerBound1, lowerBound2)
```

### ReturnArrayToPool

Returns an array to the memory pool for reuse.

```vb
' Referenced in code but implementation not visible
' Usage pattern based on code references:
ReturnArrayToPool myArray
```

---

## Processing Functions

### ProcessCommissionData

Processes commission data with intelligent chunking based on available memory.

```vb
Private Sub ProcessCommissionData(ws As Worksheet, lr As Long, cOpt As Long, cAmt As Long, cRate As Long, cRIPct As Long, cYr As Long, cMon As Long, cDur As Long, cRI As Long, cComm As Long, cCommBOP As Long, cRIComm As Long, cRICommBOP As Long, useChunking As Boolean, is64Bit As Boolean)
```

**Parameters:**
- `ws`: Worksheet containing the data
- `lr`: Last row of data
- `cXxx`: Column indices for various data elements
- `useChunking`: Whether to process in chunks for large datasets
- `is64Bit`: Whether the application is running in 64-bit mode

**Example:**
```vb
ProcessCommissionData ws, lastRow, colOpt, colAmt, colRate, colRIPct, colYr, _
                      colMon, colDur, colRI, colComm, colCommBOP, colRIComm, _
                      colRICommBOP, True, Is64BitExcel()
```

### ProcessSinglePass

Processes commission data in a single pass for smaller datasets.

```vb
Private Sub ProcessSinglePass(ws As Worksheet, lr As Long, cOpt As Long, cAmt As Long, cRate As Long, cRIPct As Long, cYr As Long, cMon As Long, cDur As Long, cRI As Long, cComm As Long, cCommBOP As Long, cRIComm As Long, cRICommBOP As Long)
```

**Example:**
```vb
' For smaller datasets where chunking is not needed
ProcessSinglePass ws, lastRow, colOpt, colAmt, colRate, colRIPct, colYr, _
                 colMon, colDur, colRI, colComm, colCommBOP, colRIComm, colRICommBOP
```

### ProcessCommissionArrays

Core calculation function for commission data from arrays.

```vb
Private Function ProcessCommissionArrays(vOpt As Variant, vAmt As Variant, vRate As Variant, vRIPct As Variant, vYr As Variant, vMon As Variant, vDur As Variant, vRI As Variant, cComm As Long, cCommBOP As Long, cRIComm As Long, cRICommBOP As Long, Optional baseOffset As Long = 0) As Variant
```

**Returns:**
- 2D array containing calculated values

---

## Dictionary and Collection Functions

### GetDistinctIDs

Creates a collection of unique IDs from a worksheet column.

```vb
Private Function GetDistinctIDs(ws As Worksheet) As Collection
```

**Parameters:**
- `ws`: Worksheet containing the data

**Returns:**
- Collection of unique IDs from column A

**Example:**
```vb
Dim uniqueIDs As Collection
Set uniqueIDs = GetDistinctIDs(Sheets("Data"))
```

### LoadDefaultPatterns

Loads patterns from a DEFAULT_PATTERN worksheet into a dictionary.

```vb
Private Sub LoadDefaultPatterns()
```

**Example:**
```vb
' Call to populate g_patternDict with values from the DEFAULT_PATTERN sheet
LoadDefaultPatterns
```

---

## Logging and Error Handling

### Log

Records messages to a log with severity level.

```vb
' Referenced in code but implementation not visible
' Usage pattern based on code references:
Log "Message text", "SEVERITY" ' Where SEVERITY might be INFO, WARNING, ERROR, etc.
```

**Example:**
```vb
Log "Processing started for sheet " & ws.Name, "INFO"
```

### HandleError

Centralized error handler to process and log errors.

```vb
' Referenced in code but implementation not visible
' Usage pattern based on code references:
HandleError "ModuleName", "FunctionName", "Error description"
```

**Example:**
```vb
HandleError "MainModule", "ProcessData", "Failed to process row " & rowNumber
```

---

## Best Practices

1. **Array Access**:
   - Always use NzLong, NzDouble, and NzString when accessing array elements to avoid runtime errors.
   - Use SafeArray2D to ensure arrays are properly dimensioned.

2. **Memory Management**:
   - Use GetPooledArray and ReturnArrayToPool for large arrays to reduce memory fragmentation.
   - Process large datasets in chunks with ProcessCommissionData using useChunking=True.

3. **Column Management**:
   - Use AddColumn to safely add new columns with automatic positioning.
   - Use GetColumnIndex to find column positions rather than hard-coding indexes.

4. **Error Handling**:
   - Implement proper error handlers in all procedures.
   - Use Log to document errors and important state changes.

5. **Performance Optimization**:
   - Use .Value2 instead of .Value when reading from Excel ranges.
   - Use LenB for faster empty string checks: `If LenB(str) > 0 Then`
   - Minimize worksheet interactions by reading data into arrays.
   - Cache array bounds in local variables rather than repeatedly calling UBound/LBound.