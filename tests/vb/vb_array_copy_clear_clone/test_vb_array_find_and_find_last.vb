' vybe-test: vb/vb_array_copy_clear_clone/test_vb_array_find_and_find_last
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
        Dim arr As Integer() = {1, 5, 8, 12, 15}
        Dim firstEven As Integer = Array.Find(arr, Function(x) x Mod 2 = 0)
        Dim lastEven As Integer = Array.FindLast(arr, Function(x) x Mod 2 = 0)
        __Check(CStr(firstEven), "8")
        __Check(CStr(lastEven), "12")
    End Sub
End Module
