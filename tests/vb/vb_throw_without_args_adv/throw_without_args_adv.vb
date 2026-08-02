' vybe-test: vb/vb_throw_without_args_adv/throw_without_args_adv
' origin: languages/vb/tests/vb/test_vb_throw_without_args_adv.rs

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
                Throw New System.Exception("Original")
            Catch
                ' Throw without args re-throws the current exception
                Throw
            End Try
        Catch ex As System.Exception
            __Check(CStr(ex.Message), "Original")
        End Try
    End Sub
End Module
