' vybe-test: vb/vb_system_exception_types_matrix/exception_throws_can_carry_inner_exception_messages
' origin: languages/vb/tests/vb/test_vb_system_exception_types_matrix.rs

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
        Try
            Try
                Throw New Exception("root")
            Catch inner As Exception
                Throw New Exception("wrapper", inner)
            End Try
        Catch ex As Exception
            __Check(CStr(ex.InnerException IsNot Nothing), "True")
            __Check(CStr(ex.InnerException.Message), "root")
        End Try
    End Sub
End Module
