' vybe-test: vb/vb_array_find_findindex_predicates/test_vb_array_findlast_with_complex_object
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

Class Product
    Public Property Category As String
    Public Property Price As Double
    Public Sub New(c As String, p As Double)
        Category = c : Price = p
    End Sub
End Class

Module Program
    Sub Main()
        Dim prods As Product() = {
            New Product("Tech", 100),
            New Product("Food", 5),
            New Product("Tech", 500)
        }
        Dim lastTech As Product = Array.FindLast(prods, Function(p) p.Category = "Tech")
        __Check(CStr(lastTech.Price), "500")
    End Sub
End Module
