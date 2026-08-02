' vybe-test: vb/vb_interface_explicit_implementation_adv/test_vb_interface_property_explicit_implementation
' origin: languages/vb/tests/vb/test_vb_interface_explicit_implementation_adv.rs

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

Interface INamed
    Property Name As String
End Interface

Class User
    Implements INamed
    Private _userName As String

    Private Property NameProp As String Implements INamed.Name
        Get
            Return _userName
        End Get
        Set(value As String)
            _userName = value
        End Set
    End Property

    Public Sub New(name As String)
        _userName = name
    End Sub
End Class

Module Program
    Sub Main()
        Dim u As New User("Bob")
        Dim n As INamed = u
        __Check(CStr(n.Name), "Bob")
    End Sub
End Module
