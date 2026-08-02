' vybe-test: vb/vb_mustoverride_property/mustoverride_property
' origin: languages/vb/tests/vb/test_vb_mustoverride_property.rs

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

MustInherit Class Config
    Public MustOverride Property ConnectionString As String
End Class

Class AppConfig
    Inherits Config
    
    Private _conn As String = "Server=Local;"
    
    Public Overrides Property ConnectionString As String
        Get
            Return _conn
        End Get
        Set(value As String)
            _conn = value
        End Set
    End Property
End Class

Module M
    Sub Main()
        Dim c As Config = New AppConfig()
        __Check(CStr(c.ConnectionString), "Server=Local;")
    End Sub
End Module
