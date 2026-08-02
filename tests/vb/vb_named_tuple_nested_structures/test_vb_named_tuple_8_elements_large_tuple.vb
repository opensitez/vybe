' vybe-test: vb/vb_named_tuple_nested_structures/test_vb_named_tuple_8_elements_large_tuple
' origin: languages/vb/tests/vb/test_vb_named_tuple_nested_structures.rs

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
        Dim t = (A:=1, B:=2, C:=3, D:=4, E:=5, F:=6, G:=7, H:=8)
        __Check(CStr(t.A & "+" & t.H), "1+8")
    End Sub
End Module
