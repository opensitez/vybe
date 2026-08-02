' vybe-test: vb/vb_property_writeonly_set_semantics/test_vb_property_set_access_modifier_narrowing
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

Class ProtectedSetProp
    Public Property Title As String
        Get
            Return _title
        End Get
        Protected Set(value As String)
            _title = value
        End Set
    End Property

    Private _title As String

    Public Sub New(t As String)
        Title = t
    End Sub
End Class

Module Program
    Sub Main()
        Dim p As New ProtectedSetProp("Initial")
        __Check(CStr(p.Title), "Initial")
    End Sub
End Module
