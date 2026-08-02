' vybe-test: vb/vb_array_sort_custom_comparer/test_vb_array_sort_double_keys_string_items
' origin: languages/vb/tests/vb/test_vb_array_sort_custom_comparer.rs

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
        Dim scores As Double() = {98.5, 87.0, 92.3}
        Dim students As String() = {"Alice", "Bob", "Charlie"}
        Array.Sort(scores, students)
        __Check(CStr(students(0)), "Bob")
        __Check(CStr(students(2)), "Alice")
    End Sub
End Module
