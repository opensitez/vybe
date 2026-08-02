' vybe-test: vb/vb_spec_object_model/object_model_spec_base_reference_can_point_to_derived_instance
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

Class Animal
    Public Overridable Function Speak() As String
        Return "animal"
    End Function
End Class
Class Dog
    Inherits Animal
    Public Overrides Function Speak() As String
        Return "dog"
    End Function
End Class
Module M
    Sub Main()
        Dim pet As Animal = New Dog()
        __Check(CStr(pet.Speak()), "dog")
    End Sub
End Module
