' vybe-test: vb/vb_linq_all_any_predicates/test_vb_linq_contains_complex_object_reference
' origin: languages/vb/tests/vb/test_vb_linq_all_any_predicates.rs

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

Imports System.Linq

Class Product
    Public Property Name As String
    Public Sub New(n As String) : Name = n : End Sub
End Class

Module Program
    Sub Main()
        Dim p1 As New Product("Laptop")
        Dim p2 As New Product("Phone")
        Dim list = {p1}
        __Check(CStr(list.Contains(p1) & "|" & list.Contains(p2)), "True|False")
    End Sub
End Module
