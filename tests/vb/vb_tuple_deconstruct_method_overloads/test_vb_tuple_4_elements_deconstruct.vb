' vybe-test: vb/vb_tuple_deconstruct_method_overloads/test_vb_tuple_4_elements_deconstruct
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
        Dim t = (1, 2, 3, 4)
        Dim a As Integer = 0, b As Integer = 0, c As Integer = 0, d As Integer = 0
        t.Deconstruct(a, b, c, d)
        __Check(CStr(a + b + c + d), "10")
    End Sub
End Module
