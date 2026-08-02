' vybe-test: vb/vb_declare_auto_function/declare_auto_function
' origin: languages/vb/tests/vb/test_vb_declare_auto_function.rs

' Vybe test harness — Visual Basic.
'
' Real VB source alongside harness/go/check.go and harness/js/check.js, the way
' test262's assert.js is JavaScript.
'
' A test's verdict is its EXIT CODE. __Check prints its diagnostic BEFORE
' throwing: an uncaught exception surfaces as `RuntimeError: [object]`, which
' says nothing at all.

Module VybeCheck
    Sub __Check(got As String, want As String)
        If got <> want Then
            Console.WriteLine("FAIL: want [" & want & "] got [" & got & "]")
            Throw New Exception("assertion failed")
        End If
    End Sub
End Module

Module M
    ' Declare statement is used to call external DLLs
    ' We just test the syntax parsing here since we can't guarantee User32 is available in test environment
    Declare Auto Function MessageBox Lib "user32.dll" (ByVal hWnd As Integer, ByVal txt As String, ByVal caption As String, ByVal Typ As Integer) As Integer
    
    Sub Main()
        Dim parsed As Boolean = True
        __Check(CStr(parsed), "True")
    End Sub
End Module
