' vybe-test: vb/vb_readonly_fields_props/readonly_fields_in_constructors
' origin: languages/vb/tests/vb/test_vb_readonly_fields_props.rs

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

Class Config
    Public ReadOnly BaseUrl As String
    
    Public Sub New(url As String)
        BaseUrl = url
    End Sub
    
    ' VB.NET allows assigning to ReadOnly fields in any constructor
    Public Sub New()
        Me.New("http://localhost")
    End Sub
End Class

Module M
    Sub Main()
        Dim c1 As New Config()
        __Check(CStr(c1.BaseUrl), "http://localhost")
        
        Dim c2 As New Config("http://test")
        __Check(CStr(c2.BaseUrl), "http://test")
    End Sub
End Module
