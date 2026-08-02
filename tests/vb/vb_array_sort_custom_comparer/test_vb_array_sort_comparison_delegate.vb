' vybe-test: vb/vb_array_sort_custom_comparer/test_vb_array_sort_comparison_delegate
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
        Dim arr As Integer() = {1, 5, 2, 9, 3}
        Array.Sort(arr, Function(x, y) y.CompareTo(x))
        __Check(CStr(String.Join(",", arr)), "9,5,3,2,1")
    End Sub
End Module
