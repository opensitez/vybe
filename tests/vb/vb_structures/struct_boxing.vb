' vybe-test: vb/vb_structures/struct_boxing
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
Dim s1 As S
s1.V = 1
Dim obj As Object = s1
Dim s2 As S = DirectCast(obj, S)
__Check(CStr(s2.V), "1")
End Sub
End Module
