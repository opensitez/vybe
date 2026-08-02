' vybe-test: vb/vb_tuple_deconstruct_method_overloads/test_vb_tuple_comparison_operators
' origin: languages/vb/tests/vb/test_vb_tuple_deconstruct_method_overloads.rs

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
        Dim t1 = (1, 10)
        Dim t2 = (1, 20)
        Dim t3 = (2, 5)
        __Check(CStr((t1 < t2) & "|" & (t2 < t3)), "True|True")
    End Sub
End Module
