' vybe-test: vb/vb_array_copy_clear_clone/test_vb_array_find_all_matches
' origin: languages/vb/tests/vb/test_vb_array_copy_clear_clone.rs

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
        Dim arr As Integer() = {1, 2, 3, 4, 5, 6}
        Dim evens As Integer() = Array.FindAll(arr, Function(x) x Mod 2 = 0)
        __Check(CStr(String.Join(",", evens)), "2,4,6")
    End Sub
End Module
