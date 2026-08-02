' vybe-test: vb/vb_callbyname_function_invocation/test_vb_callbyname_optional_arguments_omitted
' origin: languages/vb/tests/vb/test_vb_callbyname_function_invocation.rs

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

Imports Microsoft.VisualBasic

Module Program
    Class Config
        Public Function Build(host As String, Optional port As Integer = 80) As String
            Return host & ":" & port
        End Function
    End Class

    Sub Main()
        Dim c As New Config()
        Dim res = CallByName(c, "Build", CallType.Method, "localhost")
        __Check(CStr(res), "localhost:80")
    End Sub
End Module
