' vybe-test: vb/vb_array_copy_clear_clone/test_vb_array_fill_value
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
        Dim arr(4) As Integer
        Array.Fill(arr, 42)
        __Check(CStr(String.Join(",", arr)), "42,42,42,42,42")
    End Sub
End Module
