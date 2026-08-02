' vybe-test: vb/vb_array_find_findindex_predicates/test_vb_array_find_with_complex_object_predicate
' origin: languages/vb/tests/vb/test_vb_array_find_findindex_predicates.rs

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

Class Person
    Public Property Name As String
    Public Property Age As Integer
    Public Sub New(n As String, a As Integer)
        Name = n : Age = a
    End Sub
End Class

Module Program
    Sub Main()
        Dim people As Person() = {New Person("Alice", 25), New Person("Bob", 35), New Person("Charlie", 30)}
        Dim found As Person = Array.Find(people, Function(p) p.Age > 30)
        __Check(CStr(found.Name), "Bob")
    End Sub
End Module
