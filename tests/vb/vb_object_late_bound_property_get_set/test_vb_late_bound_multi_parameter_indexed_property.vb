' vybe-test: vb/vb_object_late_bound_property_get_set/test_vb_late_bound_multi_parameter_indexed_property
' origin: languages/vb/tests/vb/test_vb_object_late_bound_property_get_set.rs

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
    Class Matrix2D
        Private data(1, 1) As Double
        Default Public Property Item(row As Integer, col As Integer) As Double
            Get
                Return data(row, col)
            End Get
            Set(value As Double)
                data(row, col) = value
            End Set
        End Property
    End Class

    Sub Main()
        Dim obj As Object = New Matrix2D()
        obj(0, 1) = 7.7
        Dim val As Double = CDbl(obj(0, 1))
        __Check(CStr(val), "7.7")
    End Sub
End Module
