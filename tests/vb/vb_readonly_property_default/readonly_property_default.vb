' vybe-test: vb/vb_readonly_property_default/readonly_property_default
' origin: languages/vb/tests/vb/test_vb_readonly_property_default.rs

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

Class Configuration
    ' ReadOnly auto-property with default initializer
    Public ReadOnly Property Version As String = "1.0.0"
    
    ' Property with backing field
    Private _maxRetries As Integer = 3
    Public ReadOnly Property MaxRetries As Integer
        Get
            Return _maxRetries
        End Get
    End Property
End Class

Module M
    Sub Main()
        Dim c As New Configuration()
        __Check(CStr(c.Version), "1.0.0")
        __Check(CStr(c.MaxRetries), "3")
    End Sub
End Module
