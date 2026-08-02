' vybe-test: vb/generators/iterator_function_returns_continuation
' origin: languages/vb/tests/vb/test_generators.rs

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
    Function Count()
        Yield 1
        Yield 2
    End Function

    Sub Main()
        __Check(CStr(Count()), "[continuation]")
    End Sub
End Module
