' vybe-test: vb/vb_property_parameters/property_parameters_indexed
' origin: languages/vb/tests/vb/test_vb_property_parameters.rs

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

Class Cache
    Private data(10) As String
    
    ' Property with parameters (Indexed Property)
    Public Property ItemAt(index As Integer) As String
        Get
            Return data(index)
        End Get
        Set(value As String)
            data(index) = value
        End Set
    End Property
End Class

Module M
    Sub Main()
        Dim c As New Cache()
        c.ItemAt(5) = "Stored"
        __Check(CStr(c.ItemAt(5)), "Stored")
    End Sub
End Module
