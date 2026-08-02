' vybe-test: vb/vb_redim_statement/redim_multidimensional
' origin: languages/vb/tests/vb/test_vb_redim_statement.rs

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
        Dim grid(,) As Integer
        ReDim grid(1, 1)
        grid(0, 0) = 1
        grid(1, 1) = 5
        __Check(CStr(grid(1, 1)), "5")
        
        ' ReDim Preserve can only change the last dimension
        ReDim Preserve grid(1, 2)
        grid(1, 2) = 10
        __Check(CStr(grid(1, 2)), "10")
        __Check(CStr(grid(0, 0)), "1") ' Still 1
    End Sub
End Module
