' vybe-test: vb/vb_callbyname_function_invocation/test_vb_callbyname_overloaded_method_resolution
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
    Class Formatter
        Public Function Format(val As Integer) As String
            Return "Int:" & val
        End Function
        Public Function Format(val As String) As String
            Return "Str:" & val
        End Function
    End Class

    Sub Main()
        Dim f As New Formatter()
        Dim res1 = CallByName(f, "Format", CallType.Method, 99)
        Dim res2 = CallByName(f, "Format", CallType.Method, "Text")
        __Check(CStr(res1 & "|" & res2), "Int:99|Str:Text")
    End Sub
End Module
