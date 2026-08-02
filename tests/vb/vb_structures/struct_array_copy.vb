' vybe-test: vb/vb_structures/struct_array_copy
' origin: languages/vb/tests/vb/test_vb_structures.rs

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

Structure S
Public V As Integer
End Structure
Module M
Sub Main()
Dim arr(1) As S
arr(0).V = 5
Dim s1 = arr(0)
s1.V = 10
__Check(CStr(arr(0).V), "5")
End Sub
End Module
