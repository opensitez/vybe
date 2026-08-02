' vybe-test: vb/vb_property_writeonly_set_semantics/test_vb_property_writeonly_password_setter
' origin: languages/vb/tests/vb/test_vb_property_writeonly_set_semantics.rs

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

Class UserAccount
    Private _passwordHash As String

    Public WriteOnly Property Password As String
        Set(value As String)
            _passwordHash = "HASH_" & value
        End Set
    End Property

    Public Function CheckPassword(input As String) As Boolean
        Return _passwordHash = "HASH_" & input
    End Function
End Class

Module Program
    Sub Main()
        Dim acc As New UserAccount()
        acc.Password = "Secret123"
        __Check(CStr(acc.CheckPassword("Secret123")), "True")
        __Check(CStr(acc.CheckPassword("Wrong")), "False")
    End Sub
End Module
