VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClSelectQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

                    '**********************************************
                    '***              quanTest                  ***
                    '***    Programmer : AYON BANERJEE          ***
                    '***        Mobile : 098303 55**3           ***
                    '***    Email : y.ayon@yahoo.co.in          ***
                    '***  Created : 01/08/2006 AT 01:28 PM      ***
                    '**********************************************




Option Explicit
    '''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' This is the Select Query Class
    '''''''''''''''''''''''''''''''''''''''''''''''''''''

    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' ADODB Connection variables
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim cn As New ADODB.Connection
    Public rs As New ADODB.Recordset
    
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Prepare connectionString
    '''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim connstring1 As String
    Dim connstring2 As String
    Dim fullconnstring As String
   
   
Function SelectQuery(ByVal query As String) As Boolean
       
    SelectQuery = False
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' function to retrieve values from database
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    On Error GoTo errorHandler
    
    connstring1 = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="
    connstring2 = "\dbQuan.mdb;Persist Security Info=False;Jet OLEDB:Database Password=iamnot"
    fullconnstring = connstring1 + App.Path + connstring2
    
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Open Connection
    '''''''''''''''''''''''''''''''''''''''''''''''''''''
    If cn.State = 0 Then
        cn.Open (fullconnstring)
    Else
        ' NOT SET YET
    End If
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Execute Query on the open connection
    '''''''''''''''''''''''''''''''''''''''''''''''''''''
    If cn.State = 1 Then
        rs.Open query, cn
        
    Else
        'storeManagerMain.stmtDebug = "opening recordset"
    End If
    
    GoTo skipErrorHandler

errorHandler:

    ' error: database is not present in the default
    If Err.Number = -2147467259 Then
        MsgBox "Trying to recover from error... Please restart Application...", vbCritical + vbOKOnly
        Logger ("DEBUG: " & Err.Number & vbCrLf & "DESCRIPTION: " & Err.Description & vbCrLf & "SOURCE: " & Err.Source)
    
    ' other unhandled errors
    Else
        MsgBox "error in SelectQuery classModule :: Programmer", vbExclamation
        Logger ("DEBUG: " & Err.Number & vbCrLf & "DESCRIPTION: " & Err.Description & vbCrLf & "SOURCE: " & Err.Source)

    End If
Exit Function

skipErrorHandler:
    SelectQuery = True
End Function

Function closeSelectQuery()
    '''''''''''''''''''''''''''''''''''''''''''''''''
    ' Function to close recordset and connection
    ' variables if they are open
    '''''''''''''''''''''''''''''''''''''''''''''''''
    If rs.State = 1 Then
        rs.Close
    End If
    If cn.State = 1 Then
        cn.Close
    End If
End Function
