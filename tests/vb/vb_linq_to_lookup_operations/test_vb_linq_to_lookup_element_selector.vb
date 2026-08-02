' vybe-test: vb/vb_linq_to_lookup_operations/test_vb_linq_to_lookup_element_selector
' origin: languages/vb/tests/vb/test_vb_linq_to_lookup_operations.rs

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
        Dim items = {
            New With {.Cat = "A", .Val = 10},
            New With {.Cat = "A", .Val = 20},
            New With {.Cat = "B", .Val = 30}
        }

        Dim lookup = items.ToLookup(Function(i) i.Cat, Function(i) i.Val)
        __Check(CStr(String.Join(",", lookup("A"))), "10,20")
    End Sub
End Module
