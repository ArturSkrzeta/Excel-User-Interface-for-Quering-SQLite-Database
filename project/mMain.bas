Attribute VB_Name = "mMain"
Option Explicit

Sub Main()

    Dim conn As Object:         Set conn = CreateObject("ADODB.Connection")
    Dim rst As Object:          Set rst = CreateObject("ADODB.Recordset")

    On Error GoTo CloseADODBConnection
    Application.EnableEvents = False

    Call OpenSession(conn, rst)
    Call ExecuteQuery(conn, rst)
    
CloseADODBConnection:
    
    Application.EnableEvents = True
    Call KillSession(conn, rst)

End Sub


Private Sub ClearOldData()

    ActiveSheet.Range("B9").CurrentRegion.ClearContents

End Sub


Private Sub OpenSession(conn As Object, rst As Object)

    Dim dbPath As String:       dbPath = "C:\Users\U742905\Lufthansa Group\Procurement Tower - Outlook\data.db"
    Dim openStr As String:      openStr = "DRIVER=SQLite3 ODBC Driver;Database=" & dbPath & ";"
    conn.Open openStr

End Sub


Private Sub ExecuteQuery(conn As Object, rst As Object)

    Dim i As Integer
    Dim strSQL As String
    Dim callName As String:         callName = ActiveSheet.Shapes(Application.Caller).TextFrame.Characters.Text
    Dim timeOfUpdate As Date:       timeOfUpdate = Now
    Dim rng As Range:               Set rng = Selection.Resize(1, 1)
    rng.Select
    
    If callName = "Show" Then
    
        Call ClearOldData
        
        If ActiveSheet.optAll Then
            strSQL = "SELECT * FROM transactions WHERE deleted = 'False';"
        ElseIf ActiveSheet.optVerified Then
            strSQL = "SELECT * FROM transactions WHERE verified = 'True' AND deleted = 'False';"
        ElseIf ActiveSheet.optNotVerified Then
            strSQL = "SELECT * FROM transactions WHERE verified = 'False' AND deleted = 'False';"
        End If
        
        rst.Open strSQL, conn
    
        With rst
            For i = 1 To .Fields.Count
                ActiveSheet.Cells(9, i + 1) = .Fields(i - 1).Name
            Next i
        End With
    
        ActiveSheet.Range("B10").CopyFromRecordset rst
        
        ActiveSheet.Columns("F:F").NumberFormat = "dd/mm/yyyy hh:mm:ss"
        
    ElseIf callName = "Set to verified" Then
        
        strSQL = "UPDATE transactions SET verified = 'True', last_update = '" & Environ("username") & " - " & timeOfUpdate & "' WHERE transaction_id = '" & Selection.Value & "';"
        rst.Open strSQL, conn
        Selection.Offset(0, 4).Value = "True "
       
        Selection.Offset(0, 6).Value = Environ("username") & " - " & timeOfUpdate
        
    ElseIf callName = "Get Payer" Then
    
        strSQL = "SELECT p.name, p.country " & _
                    "FROM payers p " & _
                    "JOIN transactions t " & _
                    "ON p.id = t.payer_id " & _
                    "WHERE t.transaction_id = '" & Selection.Value & "';"
                    
        rst.Open strSQL, conn
        
        With rst
            For i = 1 To .Fields.Count
                ActiveSheet.Cells(4, i + 7) = .Fields(i - 1).Name
            Next i
        End With
        
        ActiveSheet.Range("H5").CopyFromRecordset rst
        
    ElseIf callName = "Delete" Then
    
        Dim vbAnswer As Integer:        vbAnswer = MsgBox("Are you want to delete that transaction row?", vbQuestion + vbYesNo, "Deletion")
        
        If Not vbAnswer = vbYes Then Exit Sub
    
        strSQL = "UPDATE transactions SET deleted = 'True', last_update = '" & Environ("username") & " - " & timeOfUpdate & "' WHERE transaction_id = '" & Selection.Value & "';"
        
        rst.Open strSQL, conn
        
        Selection.Offset(0, 5).Value = "True "
       
        Selection.Offset(0, 6).Value = Environ("username") & " - " & timeOfUpdate
        
    End If
    
End Sub


Private Sub KillSession(conn As Object, rst As Object)
    
    On Error Resume Next
    rst.Close
    conn.Close
    Set rst = Nothing
    Set conn = Nothing

End Sub
