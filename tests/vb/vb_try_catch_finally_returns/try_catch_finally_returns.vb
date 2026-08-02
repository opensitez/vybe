' vybe-test: vb/vb_try_catch_finally_returns/try_catch_finally_returns
' origin: languages/vb/tests/vb/test_vb_try_catch_finally_returns.rs

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
    Function TestReturn() As Integer
        Try
            Return 1
        Catch
            Return 2
        Finally
            ' Cannot Return from Finally in VB.NET, but can write to console
            __Check(CStr("Finally"), "Finally")
        End Try
    End Function

    Function TestThrow() As Integer
        Try
            Throw New Exception("Error")
        Catch
            Return 3
        Finally
            __Check(CStr("Finally2"), "1")
        End Try
    End Function

    Sub Main()
        __Check(CStr(TestReturn()), "Finally2")
        __Check(CStr(TestThrow()), "3")
    End Sub
End Module
