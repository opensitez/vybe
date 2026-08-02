' vybe-test: vb/vb_info_lbound_ubound/info_lbound_ubound
' origin: languages/vb/tests/vb/test_vb_info_lbound_ubound.rs

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
        ' Array bounds functions
        Dim arr(2) As Integer
        
        __Check(CStr(LBound(arr)), "0") ' Usually 0
        __Check(CStr(UBound(arr)), "2") ' 2
        
        ' Multi-dimensional bounds
        Dim matrix(3, 4) As Integer
        __Check(CStr(UBound(matrix, 1)), "3") ' Rank 1 -> 3
        __Check(CStr(UBound(matrix, 2)), "4") ' Rank 2 -> 4
    End Sub
End Module
