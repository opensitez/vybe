' vybe-test: vb/vb_array_convertall_transformations/test_vb_array_convertall_int_to_string
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

Module Program
    Sub Main()
        Dim numbers As Integer() = {1, 2, 3, 4}
        Dim strings As String() = Array.ConvertAll(numbers, Function(n) "Item_" & n)
        __Check(CStr(String.Join(",", strings)), "Item_1,Item_2,Item_3,Item_4")
    End Sub
End Module
