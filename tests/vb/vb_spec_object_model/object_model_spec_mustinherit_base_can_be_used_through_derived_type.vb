' vybe-test: vb/vb_spec_object_model/object_model_spec_mustinherit_base_can_be_used_through_derived_type
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

MustInherit Class Shape
    Public MustOverride Function Name() As String
End Class
Class Circle
    Inherits Shape
    Public Overrides Function Name() As String
        Return "circle"
    End Function
End Class
Module M
    Sub Main()
        Dim value As Shape = New Circle()
        __Check(CStr(value.Name()), "circle")
    End Sub
End Module
