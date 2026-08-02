' vybe-test: vb/vb_property_indexed_multi_parameter/test_vb_indexed_property_overloading_parameter_count
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

Class GridContainer
    Default Public Property Item(x As Integer) As String
        Get
            Return "1D:" & x
        End Get
        Set(value As String)
        End Set
    End Property

    Default Public Property Item(x As Integer, y As Integer) As String
        Get
            Return "2D:" & x & "," & y
        End Get
        Set(value As String)
        End Set
    End Property
End Class

Module Program
    Sub Main()
        Dim g As New GridContainer()
        __Check(CStr(g(5) & "|" & g(5, 10)), "1D:5|2D:5,10")
    End Sub
End Module
