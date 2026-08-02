' vybe-test: vb/vb_structures/struct_inheritance_fails
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

Structure S1
End Structure
' Structure S2: Inherits S1 ' Structures cannot inherit
Module M
Sub Main()
__Check(CStr("Parsed"), "Parsed")
End Sub
End Module
