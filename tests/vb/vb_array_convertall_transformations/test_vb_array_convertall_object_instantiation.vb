' vybe-test: vb/vb_array_convertall_transformations/test_vb_array_convertall_object_instantiation
' origin: languages/vb/tests/vb/test_vb_array_convertall_transformations.rs

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
    Public ReadOnly Property Name As String
    Public Sub New(name As String)
        Me.Name = name
    End Sub
End Class

Module Program
    Sub Main()
        Dim names As String() = {"Alice", "Bob"}
        Dim users As User() = Array.ConvertAll(names, Function(n) New User(n))
        __Check(CStr(users(0).Name & "&" & users(1).Name), "Alice&Bob")
    End Sub
End Module
