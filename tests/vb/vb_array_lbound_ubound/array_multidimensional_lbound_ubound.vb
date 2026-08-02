' vybe-test: vb/vb_array_lbound_ubound/array_multidimensional_lbound_ubound
' origin: languages/vb/tests/vb/test_vb_array_lbound_ubound.rs

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
        ' Explicit bounds 1 To 3 for dimension 1, and 0 To 5 for dimension 2
        ' VB supports non-zero lower bounds in declarations
        Dim grid(1 To 3, 0 To 5) As Integer
        
        ' Dimension is 1-based index in LBound/UBound
        __Check(CStr(LBound(grid, 1)), "1")
        __Check(CStr(UBound(grid, 1)), "3")
        
        __Check(CStr(LBound(grid, 2)), "0")
        __Check(CStr(UBound(grid, 2)), "5")
    End Sub
End Module
