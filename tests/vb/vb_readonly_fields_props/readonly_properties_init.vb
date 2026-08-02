' vybe-test: vb/vb_readonly_fields_props/readonly_properties_init
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

Class User
    ' ReadOnly auto-property can be initialized at declaration
    Public ReadOnly Property Id As Integer = 100
    Public ReadOnly Property Name As String
    
    Public Sub New(name As String)
        ' Can also be initialized in constructor
        Me.Name = name
    End Sub
End Class

Module M
    Sub Main()
        Dim u As New User("Alice")
        __Check(CStr(u.Id), "100")
        __Check(CStr(u.Name), "Alice")
    End Sub
End Module
