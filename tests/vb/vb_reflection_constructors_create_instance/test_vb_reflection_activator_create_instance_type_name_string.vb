' vybe-test: vb/vb_reflection_constructors_create_instance/test_vb_reflection_activator_create_instance_type_name_string
' origin: languages/vb/tests/vb/test_vb_reflection_constructors_create_instance.rs

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

Imports System

Class Target
    Public Function Echo() As String : Return "EchoTarget" : End Function
End Class

Module Program
    Sub Main()
        Dim handle = Activator.CreateInstance(Nothing, "Target")
        Dim obj As Target = CType(handle.Unwrap(), Target)
        __Check(CStr(obj.Echo()), "EchoTarget")
    End Sub
End Module
