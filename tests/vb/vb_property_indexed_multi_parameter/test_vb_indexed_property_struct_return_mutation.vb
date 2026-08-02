' vybe-test: vb/vb_property_indexed_multi_parameter/test_vb_indexed_property_struct_return_mutation
' origin: languages/vb/tests/vb/test_vb_property_indexed_multi_parameter.rs

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

Structure Point
    Public X As Integer
    Public Y As Integer
End Structure

Class PointGrid
    Private points(1) As Point
    Default Public Property Item(idx As Integer) As Point
        Get
            Return points(idx)
        End Get
        Set(value As Point)
            points(idx) = value
        End Set
    End Property
End Class

Module Program
    Sub Main()
        Dim pg As New PointGrid()
        pg(0) = New Point With {.X = 10, .Y = 20}
        __Check(CStr(pg(0).X & "," & pg(0).Y), "10,20")
    End Sub
End Module
