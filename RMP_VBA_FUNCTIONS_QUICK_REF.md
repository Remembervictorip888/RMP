# RMP Excel VBA Utility Functions Quick Reference

## Data Handling
- **NzLong(data, row, col)** - Safe conversion to Long from array element
- **NzDouble(data, row, col)** - Safe conversion to Double from array element
- **NzString(data, row, col)** - Safe conversion to String from array element
- **SafeArray2D(inputValue)** - Ensures variant is a valid 2D array
- **ParseDate(value, outDate)** - Handles various date formats including YYYYMMDD
- **TryConvertVariant(value, result, "type")** - Safe type conversion with error handling

## Column Management
- **AddColumn(ws, colName)** - Adds column header with automatic positioning
- **GetColumnIndex(ws, colName)** - Finds column index by name (0 if not found)
- **RefreshHeaderMap(ws)** - Updates internal header mapping for faster lookups

## Memory & Performance
- **GetPooledArray(rows, cols, lb1, lb2)** - Gets array from memory pool
- **ReturnArrayToPool(array)** - Returns array to memory pool
- **GetAvailableMemoryMB()** - Estimates available system memory

## Data Processing
- **GetDistinctIDs(ws)** - Gets collection of unique IDs from column A
- **LoadDefaultPatterns()** - Loads patterns from DEFAULT_PATTERN sheet
- **ProcessCommissionData(...)** - Processes data with chunking for large datasets

## Logging & Error Handling
- **Log(message, severity)** - Records message with severity level
- **HandleError(module, function, description)** - Centralized error handler

## Best Practices

1. **Safe Array Access**: Always use Nz* functions (NzLong, NzDouble, NzString) when accessing array elements.

2. **Memory Optimization**:
   - Use .Value2 instead of .Value for ranges
   - Use LenB for faster empty string checks: `If LenB(str) > 0 Then`
   - Process large datasets in chunks

3. **Column Management**:
   - Use AddColumn/GetColumnIndex instead of hard-coding indices
   - Call RefreshHeaderMap before column operations

4. **Error Handling**:
   - Always implement On Error handlers
   - Use Log to record important events and errors

5. **Array Optimization**:
   - Cache array bounds: `lb = LBound(arr); ub = UBound(arr)`
   - Use SafeArray2D to normalize array dimensions
   - For large arrays, use pooled memory functions