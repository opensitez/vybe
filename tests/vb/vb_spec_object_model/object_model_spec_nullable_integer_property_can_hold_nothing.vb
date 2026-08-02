' vybe-test: vb/vb_spec_object_model/object_model_spec_nullable_integer_property_can_hold_nothing
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

Class Holder
    Public Property Value As Integer?
End Class
Module M
    Sub Main()
        Dim h As New Holder()
        h.Value = Nothing
        __Check(CStr(IsNothing(h.Value)), "True")
    End Sub
End Module
