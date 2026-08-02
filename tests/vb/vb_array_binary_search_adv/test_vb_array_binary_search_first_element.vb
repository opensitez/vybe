' vybe-test: vb/vb_array_binary_search_adv/test_vb_array_binary_search_first_element
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
        Dim arr As Integer() = {100, 200, 300}
        Dim idx As Integer = Array.BinarySearch(arr, 100)
        __Check(CStr(idx), "0")
    End Sub
End Module
