' vybe-test: vb/vb_arrays_bounds/array_initialization_with_bounds
' origin: languages/vb/tests/vb/test_vb_arrays_bounds.rs

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
        ' If bounds are provided and initialized, they must match the initializers length
        ' Dim arr(2) As Integer = {1, 2, 3} ' 3 elements (0,1,2)
        Dim arr(2) As Integer = {10, 20, 30}
        __Check(CStr(arr(2)), "30")
    End Sub
End Module
