VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MySQLConnector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private Const MySQLDriver As String = "Driver=MySQL ODBC 5.3 UNICODE Driver;"

Private Sub Class_Initialize()
    
End Sub


Public Function CanConnectToTestDatabase() As Boolean
'this is just to test to validate a connection

    Dim TestConnection As ADODB.Connection
    
On Error GoTo Except
Try:
    Set TestConnection = ConnectToDatabase("localhost", "dummy_connection", "root", "")
    CanConnectToTestDatabase = True
    TestConnection.Close
    GoTo Finally
Except:
    CanConnectToTestDatabase = False
    MsgBox Err.Number & vbNewLine & Err.Source & vbNewLine & Err.Description
    GoTo Finally
Finally:
    Err.Clear
    Set TestConnection = Nothing

End Function

Public Function ConnectToTestDatabase() As ADODB.Connection
' get a connection object for further use

    Dim conn As ADODB.Connection
    Set ConnectToTestDatabase = ConnectToDatabase("localhost", "dummy_connection", "root", "")

End Function

Public Function ConnectToDatabase(ByVal Server As String, _
                                    ByVal DBName As String, _
                                    ByVal UID As String, _
                                    ByVal PWD As String) As ADODB.Connection

'input: connection parameters
'return: a active connection object

    Dim conn As ADODB.Connection
    Set conn = New ADODB.Connection
    Dim ConnStr As String
    
    ConnStr = MySQLDriver & _
            "Server=" & Server & ";" & _
            "Database=" & DBName & ";" & _
            "UID=" & UID & ";" & _
            "PWD=" & PWD & ";" & _
            "option=3"
    ' set provider and connect string, then open the connection
    conn.Provider = "MSDASQL.1"
    conn.ConnectionString = ConnStr
    conn.Open
    'pop some msgbox
'    MsgBox "connection successful"
    
    Set ConnectToDatabase = conn
End Function

Public Sub CreateAndUseDatabase(ByVal DBName As String, ActiveConn As ADODB.Connection)
' this sub checks first whether there's already existing database named after given dbname
' pay attention to space when typing a query

    ActiveConn.Execute "CREATE DATABASE IF NOT EXISTS " & DBName & ";"
    ActiveConn.Execute "USE " & DBName & ";"
    MsgBox "Database changed to " & DBName

End Sub


