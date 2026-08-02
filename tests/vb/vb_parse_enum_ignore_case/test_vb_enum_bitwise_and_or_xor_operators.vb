' vybe-test: vb/vb_parse_enum_ignore_case/test_vb_enum_bitwise_and_or_xor_operators
' origin: languages/vb/tests/vb/test_vb_parse_enum_ignore_case.rs

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

<Flags>
Enum Options
    OptA = 1
    OptB = 2
    OptC = 4
End Enum

Module Program
    Sub Main()
        Dim combined = Options.OptA Or Options.OptB
        Dim toggled = combined Xor Options.OptA
        __Check(CStr(toggled.ToString()), "OptB")
    End Sub
End Module
