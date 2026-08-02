' vybe-test: vb/vb_linq_select_many_projections/test_vb_linq_compound_from_flatten_nested_collections
' origin: languages/vb/tests/vb/test_vb_linq_select_many_projections.rs

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

Module Program
    Sub Main()
        Dim categories = {
            New With {.Name = "Fruit", .Items = {"Apple", "Banana"}},
            New With {.Name = "Veggie", .Items = {"Carrot", "Pea"}}
        }

        Dim allItems = From cat In categories
                       From item In cat.Items
                       Select item

        __Check(CStr(String.Join(",", allItems)), "Apple,Banana,Carrot,Pea")
    End Sub
End Module
