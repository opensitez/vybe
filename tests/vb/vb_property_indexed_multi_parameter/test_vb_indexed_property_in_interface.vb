' vybe-test: vb/vb_property_indexed_multi_parameter/test_vb_indexed_property_in_interface
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

Interface IMatrix
    Default Property Item(r As Integer, c As Integer) As Double
End Interface

Class DoubleMatrix
    Implements IMatrix
    Private arr(1, 1) As Double
    Default Public Property Item(r As Integer, c As Integer) As Double Implements IMatrix.Item
        Get
            Return arr(r, c)
        End Get
        Set(value As Double)
            arr(r, c) = value
        End Set
    End Property
End Class

Module Program
    Sub Main()
        Dim m As IMatrix = New DoubleMatrix()
        m(0, 1) = 3.14
        __Check(CStr(m(0, 1)), "3.14")
    End Sub
End Module
