' vybe-test: vb/vb_enums/enum_math_operations
' origin: languages/vb/tests/vb/test_vb_enums.rs

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

Enum E
A = 1
B = 2
End Enum
Module M
Sub Main()
__Check(CStr(CInt(E.A + E.B)), "3")
End Sub
End Module
