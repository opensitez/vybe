' vybe-test: vb/vb_static_locals/static_local_is_isolated_per_function
' origin: languages/vb/tests/vb/test_vb_static_locals.rs

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
    Function LeftCounter() As Integer
        Static total As Integer = 0
        total = total + 1
        Return total
    End Function

    Function RightCounter() As Integer
        Static total As Integer = 10
        total = total + 5
        Return total
    End Function

    Sub Main()
        __Check(CStr(LeftCounter()), "1")
        __Check(CStr(RightCounter()), "15")
        __Check(CStr(LeftCounter()), "2")
        __Check(CStr(RightCounter()), "20")
    End Sub
End Module
