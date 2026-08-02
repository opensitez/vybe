' vybe-test: vb/vb_builtins/math_methods
' origin: languages/vb/tests/vb/vb_builtins_test.rs

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

Module Program
    Sub Main()
        __Check(CStr(Math.Abs(-7)), "7")
        __Check(CStr(Math.Sqrt(16)), "4")
        __Check(CStr(Math.Pow(2, 8)), "256")
        __Check(CStr(Math.Min(3, 7)), "3")
        __Check(CStr(Math.Max(3, 7)), "7")
    End Sub
End Module
