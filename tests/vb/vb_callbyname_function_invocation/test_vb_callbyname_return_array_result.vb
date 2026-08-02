' vybe-test: vb/vb_callbyname_function_invocation/test_vb_callbyname_return_array_result
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
    Class Provider
        Public Function GetTags() As String()
            Return New String() {"tag1", "tag2"}
        End Function
    End Class

    Sub Main()
        Dim p As New Provider()
        Dim tags As String() = CType(CallByName(p, "GetTags", CallType.Method), String())
        __Check(CStr(String.Join(",", tags)), "tag1,tag2")
    End Sub
End Module
