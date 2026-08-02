' vybe-test: vb/vb_interface_default_methods_adv/test_vb_interface_write_only_property
' origin: languages/vb/tests/vb/test_vb_interface_default_methods_adv.rs

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

Interface IPasswordReceiver
    WriteOnly Property Password As String
End Interface

Class Service
    Implements IPasswordReceiver
    Private _pwd As String
    Public WriteOnly Property Password As String Implements IPasswordReceiver.Password
        Set(value As String)
            _pwd = value
        End Set
    End Property
    Public Function Verify(p As String) As Boolean
        Return _pwd = p
    End Function
End Class

Module Program
    Sub Main()
        Dim s As New Service()
        Dim r As IPasswordReceiver = s
        r.Password = "Secret123"
        __Check(CStr(s.Verify("Secret123")), "True")
    End Sub
End Module
