' vybe-test: vb/vb_array_init_bounds/array_initialization_1d
' origin: languages/vb/tests/vb/test_vb_array_init_bounds.rs

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
        Dim arr1 As Integer() = {1, 2, 3}
        Dim arr2() As Integer = {4, 5, 6}
        __Check(CStr(arr1(0) + arr2(0)), "5")
    End Sub
End Module
