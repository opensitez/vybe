' vybe-test: vb/vb_string_join_enumerable_overloads/test_vb_string_join_custom_class_enumerable
' origin: languages/vb/tests/vb/test_vb_string_join_enumerable_overloads.rs

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

Imports System.Collections.Generic

Class Person
    Public Property Name As String
    Public Sub New(n As String)
        Me.Name = n
    End Sub
    Public Overrides Function ToString() As String
        Return Name
    End Function
End Class

Module Program
    Sub Main()
        Dim people As New List(Of Person) From {New Person("Alice"), New Person("Bob")}
        __Check(CStr(String.Join(" & ", people)), "Alice & Bob")
    End Sub
End Module
