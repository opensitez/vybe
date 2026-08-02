' vybe-test: vb/vb_reflection_constructors_create_instance/test_vb_reflection_constructor_is_public_is_private
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
Imports System.Reflection

Class Sample
    Public Sub New() : End Sub
    Private Sub New(x As Integer) : End Sub
End Class

Module Program
    Sub Main()
        Dim t = GetType(Sample)
        Dim pubCtor = t.GetConstructor(Type.EmptyTypes)
        Dim privCtor = t.GetConstructor(BindingFlags.Instance Or BindingFlags.NonPublic, Nothing, {GetType(Integer)}, Nothing)

        __Check(CStr(pubCtor.IsPublic & "|" & privCtor.IsPrivate), "True|True")
    End Sub
End Module
