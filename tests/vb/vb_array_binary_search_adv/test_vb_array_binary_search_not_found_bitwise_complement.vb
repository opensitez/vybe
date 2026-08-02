' vybe-test: vb/vb_array_binary_search_adv/test_vb_array_binary_search_not_found_bitwise_complement
' origin: languages/vb/tests/vb/test_vb_array_binary_search_adv.rs

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
        Dim arr As Integer() = {10, 20, 30, 40, 50}
        Dim idx As Integer = Array.BinarySearch(arr, 25)
        __Check(CStr(idx < 0), "True")
        __Check(CStr(Not idx), "2") ' Index where 25 would be inserted (2)
    End Sub
End Module
