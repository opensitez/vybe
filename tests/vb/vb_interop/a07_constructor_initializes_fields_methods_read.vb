' vybe-test: vb/vb_interop/a07_constructor_initializes_fields_methods_read
' origin: languages/vb/tests/vb/vb_interop_test.rs

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

Public Class Config
    Dim host As String
    Dim port As Integer
    Dim secure As Boolean
    Public Sub New()
        host = "localhost"
        port = 8080
        secure = True
    End Sub
    Public Function GetUrl() As String
        If secure Then
            Return "https://" & host & ":" & CStr(port)
        Else
            Return "http://" & host & ":" & CStr(port)
        End If
    End Function
End Class
Dim cfg As New Config()
__Check(CStr(cfg.GetUrl()), "https://localhost:8080")
