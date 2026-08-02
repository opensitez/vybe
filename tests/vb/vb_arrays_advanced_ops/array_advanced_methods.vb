' vybe-test: vb/vb_arrays_advanced_ops/array_advanced_methods
' origin: languages/vb/tests/vb/test_vb_arrays_advanced_ops.rs

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

Module M
    Sub Main()
        Dim arr() As Integer = {5, 2, 8, 1, 9}
        
        System.Array.Sort(arr)
        __Check(CStr(arr(0)), "1")
        __Check(CStr(arr(arr.Length - 1)), "9")
        
        System.Array.Reverse(arr)
        __Check(CStr(arr(0)), "9")
        
        Dim idx = System.Array.IndexOf(arr, 2)
        __Check(CStr(idx), "3")
    End Sub
End Module
