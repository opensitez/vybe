' vybe-test: vb/vb_indexer_multidimensional/test_vb_indexer_2d_matrix_access
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

Class SparseMatrix
    Private _data As New System.Collections.Generic.Dictionary(Of String, Double)()

    Default Public Property Item(row As Integer, col As Integer) As Double
        Get
            Dim key As String = row & "," & col
            If _data.ContainsKey(key) Then Return _data(key)
            Return 0.0
        End Get
        Set(value As Double)
            Dim key As String = row & "," & col
            _data(key) = value
        End Set
    End Property
End Class

Module Program
    Sub Main()
        Dim mat As New SparseMatrix()
        mat(1, 2) = 3.14
        __Check(CStr(mat(1, 2)), "3.14")
        __Check(CStr(mat(0, 0)), "0")
    End Sub
End Module
