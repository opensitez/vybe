' vybe-test: vb/vb_linq_skip_take_while/test_vb_linq_takewhile_complex_objects
' origin: languages/vb/tests/vb/test_vb_linq_skip_take_while.rs

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

Class Item
    Public Property Price As Double
    Public Sub New(p As Double) : Price = p : End Sub
End Class

Module Program
    Sub Main()
        Dim items = {New Item(10), New Item(20), New Item(50), New Item(5)}
        Dim cheap = items.TakeWhile(Function(i) i.Price < 30)
        __Check(CStr(cheap.Count()), "2")
    End Sub
End Module
