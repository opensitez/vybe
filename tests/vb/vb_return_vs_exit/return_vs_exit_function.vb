' vybe-test: vb/vb_return_vs_exit/return_vs_exit_function
' origin: languages/vb/tests/vb/test_vb_return_vs_exit.rs

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
    Function TestExit() As Integer
        TestExit = 10 ' Implicit return variable
        Exit Function ' Returns immediately
        TestExit = 20
    End Function

    Function TestReturn() As Integer
        Return 30 ' Explicit return
        TestReturn = 40
    End Function

    Function TestImplicit() As Integer
        TestImplicit = 50
        ' Reaches end of function, returns TestImplicit
    End Function

    Sub Main()
        __Check(CStr(TestExit()), "10")
        __Check(CStr(TestReturn()), "30")
        __Check(CStr(TestImplicit()), "50")
    End Sub
End Module
