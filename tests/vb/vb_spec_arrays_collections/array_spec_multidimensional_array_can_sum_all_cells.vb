' vybe-test: vb/vb_spec_arrays_collections/array_spec_multidimensional_array_can_sum_all_cells
' origin: languages/vb/tests/vb/test_vb_spec_arrays_collections.rs

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

Module M : Sub Main() : Dim grid(1,1) As Integer : grid(0,0)=1 : grid(0,1)=2 : grid(1,0)=3 : grid(1,1)=4 : __Check(CStr(grid(0,0)+grid(0,1)+grid(1,0)+grid(1,1)), "10") : End Sub : End Module
