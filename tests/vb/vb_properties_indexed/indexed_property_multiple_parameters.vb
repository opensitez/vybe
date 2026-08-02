' vybe-test: vb/vb_properties_indexed/indexed_property_multiple_parameters
' origin: languages/vb/tests/vb/test_vb_properties_indexed.rs

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

Class Matrix2D
    Private _data(2, 2) As Integer
    
    Public Property Item(x As Integer, y As Integer) As Integer
        Get
            Return _data(x, y)
        End Get
        Set(value As Integer)
            _data(x, y) = value
        End Set
    End Property
End Class

Module M
    Sub Main()
        Dim mat As New Matrix2D()
        mat.Item(1, 2) = 42
        __Check(CStr(mat.Item(1, 2)), "42")
    End Sub
End Module
