' vybe-test: vb/vb_reflection_constructors_create_instance/test_vb_reflection_activator_create_instance_with_args
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

Class User
    Public Property Name As String
    Public Property Age As Integer
    Public Sub New(n As String, a As Integer)
        Name = n : Age = a
    End Sub
End Class

Module Program
    Sub Main()
        Dim t = GetType(User)
        Dim u As User = CType(Activator.CreateInstance(t, "Alice", 30), User)
        __Check(CStr(u.Name & " is " & u.Age), "Alice is 30")
    End Sub
End Module
