' vybe-test: vb/vb_named_tuple_nested_structures/test_vb_named_tuple_nested_array_of_tuples
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
        Dim grid As (Row As Integer, Cols As (ColName As String, Val As Integer)())() = {
            (1, {("C1", 10), ("C2", 20)})
        }
        __Check(CStr(grid(0).Row & "->" & grid(0).Cols(0).ColName & "=" & grid(0).Cols(0).Val), "1->C1=10")
    End Sub
End Module
