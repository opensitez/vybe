' vybe-test: vb/vb_linq_of_type_filtering/test_vb_linq_of_type_filter_heterogeneous_array
' origin: languages/vb/tests/vb/test_vb_linq_of_type_filtering.rs

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

Imports System.Collections
Imports System.Linq

Module Program
    Sub Main()
        Dim mixed As Object() = {10, "Hello", 20.5, "World", 30}
        Dim strings = mixed.OfType(Of String)()
        Dim ints = mixed.OfType(Of Integer)()

        __Check(CStr(String.Join(",", strings)), "Hello,World")
        __Check(CStr(String.Join(",", ints)), "10,30")
    End Sub
End Module
