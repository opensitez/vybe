' vybe-test: vb/vb_callbyname_function_invocation/test_vb_callbyname_property_get
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
    Class User
        Public Property Username As String = "Alice"
    End Class

    Sub Main()
        Dim u As New User()
        Dim name = CallByName(u, "Username", CallType.Get)
        __Check(CStr(name), "Alice")
    End Sub
End Module
