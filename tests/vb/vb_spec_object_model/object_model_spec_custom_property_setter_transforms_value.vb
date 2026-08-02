' vybe-test: vb/vb_spec_object_model/object_model_spec_custom_property_setter_transforms_value
' origin: languages/vb/tests/vb/test_vb_spec_object_model.rs

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

Class Person
    Private _name As String
    Public WriteOnly Property LowerName As String
        Set(value As String)
            _name = LCase(value)
        End Set
    End Property
    Public Function Snapshot() As String
        Return _name
    End Function
End Class
Module M
    Sub Main()
        Dim p As New Person()
        p.LowerName = "Ada"
        __Check(CStr(p.Snapshot()), "ada")
    End Sub
End Module
