' vybe-test: vb/vb_properties_readonly/readonly_property_basic
' origin: languages/vb/tests/vb/test_vb_properties_readonly.rs

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
    Private _maxItems As Integer = 100
    
    Public ReadOnly Property MaxItems As Integer
        Get
            Return _maxItems
        End Get
    End Property
End Class

Module M
    Sub Main()
        Dim config As New Configuration()
        __Check(CStr(config.MaxItems), "100")
    End Sub
End Module
