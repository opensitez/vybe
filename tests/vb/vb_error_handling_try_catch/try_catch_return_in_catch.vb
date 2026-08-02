' vybe-test: vb/vb_error_handling_try_catch/try_catch_return_in_catch
' origin: languages/vb/tests/vb/test_vb_error_handling_try_catch.rs

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
Function F() As String
Try
Throw New System.Exception()
Catch
Return "Catch"
Finally
__Check(CStr("Finally"), "Finally")
End Try
End Function
Sub Main()
__Check(CStr(F()), "Catch")
End Sub
End Module
