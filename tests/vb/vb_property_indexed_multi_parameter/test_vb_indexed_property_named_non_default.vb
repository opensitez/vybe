' vybe-test: vb/vb_property_indexed_multi_parameter/test_vb_indexed_property_named_non_default
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

Class Table
    Private data(1, 1) As String
    Public Property ItemAt(row As Integer, col As Integer) As String
        Get
            Return data(row, col)
        End Get
        Set(value As String)
            data(row, col) = value
        End Set
    End Property
End Class

Module Program
    Sub Main()
        Dim t As New Table()
        t.ItemAt(0, 1) = "Header"
        __Check(CStr(t.ItemAt(0, 1)), "Header")
    End Sub
End Module
