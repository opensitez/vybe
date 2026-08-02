' vybe-test: vb/vb_properties_readonly/readonly_property_shadowing
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

Class BaseConfig
    Public ReadOnly Property Role As String
        Get
            Return "Guest"
        End Get
    End Property
End Class

Class AdminConfig
    Inherits BaseConfig
    Public Shadows ReadOnly Property Role As String
        Get
            Return "Admin"
        End Get
    End Property
End Class

Module M
    Sub Main()
        Dim c As New AdminConfig()
        __Check(CStr(c.Role), "Admin")
        
        Dim b As BaseConfig = c
        __Check(CStr(b.Role), "Guest")
    End Sub
End Module
