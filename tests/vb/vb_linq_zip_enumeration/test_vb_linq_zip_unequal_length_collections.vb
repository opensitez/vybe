' vybe-test: vb/vb_linq_zip_enumeration/test_vb_linq_zip_unequal_length_collections
' origin: languages/vb/tests/vb/test_vb_linq_zip_enumeration.rs

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
        Dim numbers = {1, 2, 3, 4, 5}
        Dim letters = {"A", "B"}
        Dim zipped = numbers.Zip(letters, Function(n, l) n & l)
        __Check(CStr(String.Join(",", zipped)), "1A,2B")
    End Sub
End Module
