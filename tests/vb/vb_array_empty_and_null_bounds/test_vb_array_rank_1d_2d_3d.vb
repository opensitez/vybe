' vybe-test: vb/vb_array_empty_and_null_bounds/test_vb_array_rank_1d_2d_3d
' origin: languages/vb/tests/vb/test_vb_array_empty_and_null_bounds.rs

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
        Dim a1(2) As Integer
        Dim a2(2, 2) As Integer
        Dim a3(1, 1, 1) As Integer
        __Check(CStr(a1.Rank & "," & a2.Rank & "," & a3.Rank), "1,2,3")
    End Sub
End Module
