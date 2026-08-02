' vybe-test: vb/vb_properties_default/default_property_multiple_args
' origin: languages/vb/tests/vb/test_vb_properties_default.rs

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

Class Map2D
    Private _map(5, 5) As Integer
    
    Default Public Property Cell(x As Integer, y As Integer) As Integer
        Get
            Return _map(x, y)
        End Get
        Set(value As Integer)
            _map(x, y) = value
        End Set
    End Property
End Class

Module M
    Sub Main()
        Dim map As New Map2D()
        map(2, 3) = 99
        __Check(CStr(map(2, 3)), "99")
    End Sub
End Module
