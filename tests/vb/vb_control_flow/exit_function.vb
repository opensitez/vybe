' vybe-test: vb/vb_control_flow/exit_function
' origin: languages/vb/tests/vb/test_vb_control_flow.rs

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
    Function SafeDiv(a As Integer, b As Integer) As Integer
        If b = 0 Then
            SafeDiv = -1
            Exit Function
        End If
        SafeDiv = a \ b
    End Function
    Sub Main()
        __Check(CStr(SafeDiv(10, 2)), "5")
        __Check(CStr(SafeDiv(10, 0)), "-1")
    End Sub
End Module
