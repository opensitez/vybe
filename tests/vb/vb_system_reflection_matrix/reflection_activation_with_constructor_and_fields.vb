' vybe-test: vb/vb_system_reflection_matrix/reflection_activation_with_constructor_and_fields
' origin: languages/vb/tests/vb/test_vb_system_reflection_matrix.rs

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
Imports System.Reflection

Module M
    Sub Main()
        Dim t As Type = GetType(Buildable)
        Dim obj As Object = Activator.CreateInstance(t)
        Dim ctor As ConstructorInfo = t.GetConstructor(Type.EmptyTypes)
        __Check(CStr(ctor IsNot Nothing), "True")
        Dim value As Integer = CType(t.GetField("Counter").GetValue(obj), Integer)
        __Check(CStr(value), "11")
        __Check(CStr(obj.GetType().Name), "Buildable")
    End Sub

    Class Buildable
        Public Counter As Integer = 11
    End Class
End Module
