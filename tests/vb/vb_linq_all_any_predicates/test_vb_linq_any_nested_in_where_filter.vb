' vybe-test: vb/vb_linq_all_any_predicates/test_vb_linq_any_nested_in_where_filter
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

Imports System.Collections.Generic
Imports System.Linq

Class Department
    Public Property Name As String
    Public Property Employees As List(Of String)
End Class

Module Program
    Sub Main()
        Dim depts As New List(Of Department) From {
            New Department With {.Name = "HR", .Employees = New List(Of String) From {"Alice"}},
            New Department With {.Name = "IT", .Employees = New List(Of String) From {"Bob", "Charlie"}}
        }
        Dim withCharlie = depts.Where(Function(d) d.Employees.Any(Function(e) e = "Charlie"))
        __Check(CStr(withCharlie.First().Name), "IT")
    End Sub
End Module
