' vybe-test: vb/vb_properties_readonly/readonly_property_auto_implemented
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

Class User
    Public ReadOnly Property ID As Integer = 12345
End Class

Module M
    Sub Main()
        Dim u As New User()
        __Check(CStr(u.ID), "12345")
    End Sub
End Module
