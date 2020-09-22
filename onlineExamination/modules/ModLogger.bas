Attribute VB_Name = "ModLogger"

                    '**********************************************
                    '***              quanTest                  ***
                    '***    Programmer : AYON BANERJEE          ***
                    '***        Mobile : 098303 55**3           ***
                    '***    Email : y.ayon@yahoo.co.in          ***
                    '***  Created : 01/08/2006 AT 01:28 PM      ***
                    '**********************************************



Option Explicit

Function Logger(logStmt As String)
    
    Dim filePath As String
    Dim stmt As String
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''
    '   open log code
    '''''''''''''''''''''''''''''''''''''''''''''''''''
    filePath = App.Path + "\#quanLog#" ' set path
    Open filePath For Append As #1 ' #1 is the file handle
                                    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '   write (append) into the log file
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''
    stmt = "##### Starting new session #### " + Str(Now)
    Write #1, stmt
    
    Write #1, logStmt
    'If stmtdebug <> "" Then
    '    stmtdebug = "   DEBUG: " + stmtdebug + " ErrorCode :: " + errorNo + " in form :: " + errorFormName + " event :: " + errorEvent + " description :: " + errorMsg
    '    Write #1, stmtdebug
    'Else
    '    stmtFatal = "   FATAL: ErrorCode :: " + errorNo + " in form :: " + errorFormName + " event :: " + errorEvent + " description :: " + errorMsg
    '    Write #1, stmtFatal
    'End If
    stmt = "--------------------------------------------------"
    Write #1, stmt
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '   close the log file
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Close #1
    MsgBox " An error LOG has been created! Please contact the Software Vendor Immediately", vbOKOnly + vbCritical
End Function
