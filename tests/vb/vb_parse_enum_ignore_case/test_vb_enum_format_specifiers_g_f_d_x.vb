' vybe-test: vb/vb_parse_enum_ignore_case/test_vb_enum_format_specifiers_g_f_d_x
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
Enum Modes
    Read = 1
    Write = 2
End Enum

Module Program
    Sub Main()
        Dim m = Modes.Read Or Modes.Write
        __Check(CStr([Enum].Format(GetType(Modes), m, "G") & "|" & [Enum].Format(GetType(Modes), m, "D") & "|" & [Enum].Format(GetType(Modes), m, "X")), "Read, Write|3|00000003")
    End Sub
End Module
