' vybe-test: vb/vb_spec_namespaces_modules/namespace_spec_class_inside_namespace_can_inherit_base_in_same_namespace
' origin: languages/vb/tests/vb/test_vb_spec_namespaces_modules.rs

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

Namespace Demo
    Public Class Animal
        Public Overridable Function Speak() As String
            Return "animal"
        End Function
    End Class
    Public Class Dog
        Inherits Animal
        Public Overrides Function Speak() As String
            Return "dog"
        End Function
    End Class
End Namespace
Module M
    Sub Main()
        __Check(CStr((New Demo.Dog()).Speak()), "dog")
    End Sub
End Module
