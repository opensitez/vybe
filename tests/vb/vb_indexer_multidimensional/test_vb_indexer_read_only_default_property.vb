' vybe-test: vb/vb_indexer_multidimensional/test_vb_indexer_read_only_default_property
' origin: languages/vb/tests/vb/test_vb_indexer_multidimensional.rs

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

Class ReadOnlyGrid
    Default Public ReadOnly Property Item(row As Integer, col As Integer) As String
        Get
            Return "Cell(" & row & "," & col & ")"
        End Get
    End Property
End Class

Module Program
    Sub Main()
        Dim g As New ReadOnlyGrid()
        __Check(CStr(g(2, 3)), "Cell(2,3)")
    End Sub
End Module
