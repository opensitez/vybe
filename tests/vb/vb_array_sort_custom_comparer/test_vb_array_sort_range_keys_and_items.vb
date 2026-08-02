' vybe-test: vb/vb_array_sort_custom_comparer/test_vb_array_sort_range_keys_and_items
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
        Dim keys As Integer() = {10, 40, 30, 20, 50}
        Dim vals As String() = {"A", "D", "C", "B", "E"}
        Array.Sort(keys, vals, 1, 3)
        __Check(CStr(String.Join(",", vals)), "A,B,C,D,E")
    End Sub
End Module
