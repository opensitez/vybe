' vybe-test: vb/vb_system_exception_matrix/exception_nested_try_finally_contract
' origin: languages/vb/tests/vb/test_vb_system_exception_matrix.rs

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
    Sub Main()
        Dim depth As Integer = 0
        Try
            Try
                depth = 1
            Finally
                depth += 1
            End Try
        Catch ex As Exception
            depth = -1
        End Try
        __Check(CStr(depth), "2")
    End Sub
End Module
