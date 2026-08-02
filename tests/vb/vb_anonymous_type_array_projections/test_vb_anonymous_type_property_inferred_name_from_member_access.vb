' vybe-test: vb/vb_anonymous_type_array_projections/test_vb_anonymous_type_property_inferred_name_from_member_access
' origin: languages/vb/tests/vb/test_vb_anonymous_type_array_projections.rs

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
    Public Property Username As String = "AdminUser"
End Class

Module Program
    Sub Main()
        Dim u As New User()
        ' Inferred property name .Username
        Dim obj = New With {u.Username}
        __Check(CStr(obj.Username), "AdminUser")
    End Sub
End Module
