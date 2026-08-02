' vybe-test: vb/vb_tuples_syntax/tuple_literal_unnamed
' origin: languages/vb/tests/vb/test_vb_tuples_syntax.rs

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
        Dim t = (1, "Apple")
        __Check(CStr(t.Item1), "1")
        __Check(CStr(t.Item2), "Apple")
    End Sub
End Module
