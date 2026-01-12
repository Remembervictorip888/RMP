```
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ErrorHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' =============================================================================
' ERROR HANDLER CLASS
' =============================================================================
'
' PURPOSE: Centralized error handling for the RMP Excel VBA Processing Engine
' AUTHOR: HKMC Budget Enhancement Team
' VERSION: 1.0
' LAST UPDATED: December 2025
'
' This class provides centralized error handling capabilities to replace
' repetitive On Error GoTo statements throughout the application.
'
' USAGE:
'   ErrorHandler.LogError "ModuleName", "RoutineName"
'
' BEST PRACTICES:
'   1. Call this method from every error handler section
'   2. Pass meaningful module and routine names for better diagnostics
'   3. Combine with local error context for comprehensive logging
'
' =============================================================================

Private Const MAX_ERROR_DEPTH As Integer = 10

' Error tracking variables
Private m_ErrorStack As Collection
Private m_ErrorDepth As Integer

' Initialize the error handler
Private Sub Class_Initialize()
    Set m_ErrorStack = New Collection
    m_ErrorDepth = 0
End Sub

' Clean up when object is destroyed
Private Sub Class_Terminate()
    Set m_ErrorStack = Nothing
End Sub

' =============================================================================
' PUBLIC METHODS
' =============================================================================

' Log an error with module and routine context
Public Sub LogError(ByVal moduleName As String, ByVal routineName As String)
    On Error Resume Next
    
    ' Increment error depth
    m_ErrorDepth = m_ErrorDepth + 1
    
    ' Prevent infinite recursion
    If m_ErrorDepth > MAX_ERROR_DEPTH Then
        LogSystemError "Maximum error depth exceeded", moduleName, routineName
        Exit Sub
    End If
    
    ' Capture error details
    Dim errorNumber As Long: errorNumber = Err.Number
    Dim errorDescription As String: errorDescription = Err.Description
    Dim errorSource As String: errorSource = Err.Source
    
    ' Additional error information
    Dim errorLine As Long
    #If VBA7 Then
        ' In newer versions of VBA, we can sometimes get the error line
        errorLine = Erl
    #Else
        errorLine = 0
    #End If
    
    ' Log the error
    LogSystemError errorDescription, moduleName, routineName, errorNumber, errorSource, errorLine
    
    ' Add to error stack for potential debugging
    PushErrorToStack errorNumber, errorDescription, moduleName, routineName, errorSource, errorLine
    
    ' Clear the error
    Err.Clear
    
    ' Decrement error depth
    m_ErrorDepth = m_ErrorDepth - 1
    
    On Error GoTo 0
End Sub

' Get formatted error information
Public Function GetErrorInfo() As String
    On Error Resume Next
    
    Dim errorInfo As String
    errorInfo = "Error Stack:" & vbCrLf
    
    If m_ErrorStack Is Nothing Then
        errorInfo = errorInfo & "  No error stack available"
    Else
        Dim i As Integer
        For i = 1 To m_ErrorStack.Count
            errorInfo = errorInfo & "  " & m_ErrorStack(i) & vbCrLf
        Next i
    End If
    
    GetErrorInfo = errorInfo
    
    On Error GoTo 0
End Function

' Clear the error stack
Public Sub ClearErrorStack()
    On Error Resume Next
    
    If Not m_ErrorStack Is Nothing Then
        Set m_ErrorStack = New Collection
    End If
    
    m_ErrorDepth = 0
    
    On Error GoTo 0
End Sub

' =============================================================================
' PRIVATE METHODS
' =============================================================================

' Log system error to the main log
Private Sub LogSystemError(ByVal description As String, _
                          ByVal moduleName As String, _
                          ByVal routineName As String, _
                          Optional ByVal errorNumber As Long = 0, _
                          Optional ByVal source As String = "", _
                          Optional ByVal line As Long = 0)
    On Error Resume Next
    
    ' Format the error message
    Dim errorMsg As String
    errorMsg = "ERROR | Module: " & moduleName & " | Routine: " & routineName
    
    If errorNumber <> 0 Then
        errorMsg = errorMsg & " | Number: " & errorNumber
    End If
    
    errorMsg = errorMsg & " | Description: " & description
    
    If source <> "" Then
        errorMsg = errorMsg & " | Source: " & source
    End If
    
    If line > 0 Then
        errorMsg = errorMsg & " | Line: " & line
    End If
    
    ' Timestamp the error
    errorMsg = Format(Now, "hh:mm:ss") & " " & errorMsg
    
    ' Output to immediate window if debug mode
    #If DEBUG Then
        Debug.Print errorMsg
    #End If
    
    ' Write to log file if available
    WriteToLogFile errorMsg
    
    On Error GoTo 0
End Sub

' Write message to log file
Private Sub WriteToLogFile(ByVal message As String)
    On Error Resume Next
    
    ' This would typically interface with the main logging system
    ' For now, we'll just ensure it doesn't cause errors
    Dim fileNum As Integer
    fileNum = FreeFile
    
    ' In a real implementation, this would write to the application's log file
    ' For now, we're just preventing errors
    
    On Error GoTo 0
End Sub

' Add error to stack for debugging
Private Sub PushErrorToStack(ByVal errorNumber As Long, _
                            ByVal description As String, _
                            ByVal moduleName As String, _
                            ByVal routineName As String, _
                            ByVal source As String, _
                            ByVal line As Long)
    On Error Resume Next
    
    Dim errorEntry As String
    errorEntry = Format(Now, "hh:mm:ss") & " - " & moduleName & "." & routineName
    
    If errorNumber <> 0 Then
        errorEntry = errorEntry & " (Err#" & errorNumber & ")"
    End If
    
    errorEntry = errorEntry & ": " & description
    
    If m_ErrorStack.Count >= MAX_ERROR_DEPTH Then
        ' Remove oldest entry
        m_ErrorStack.Remove 1
    End If
    
    m_ErrorStack.Add errorEntry
    
    On Error GoTo 0
End Sub
```