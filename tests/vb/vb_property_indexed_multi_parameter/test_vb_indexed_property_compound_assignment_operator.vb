' vybe-test: vb/vb_property_indexed_multi_parameter/test_vb_indexed_property_compound_assignment_operator
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

Class NumericGrid
    Private arr(2) As Integer
    Default Public Property Value(i As Integer) As Integer
        Get
            Return arr(i)
        End Get
        Set(val As Integer)
            arr(i) = val
        End Set
    End Property
End Class

Module Program
    Sub Main()
        Dim g As New NumericGrid()
        g(0) = 10
        g(0) += 5
        __Check(CStr(g(0)), "15")
    End Sub
End Module
