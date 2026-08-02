' vybe-test: vb/vb_optional_paramarrays/opt_arg_enum_default
' origin: languages/vb/tests/vb/test_vb_optional_paramarrays.rs

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
A
B
End Enum
Module M
Function F(Optional e1 As E = E.B) As String
Return e1.ToString()
End Function
Sub Main()
__Check(CStr(F()), "B")
End Sub
End Module
