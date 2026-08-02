' vybe-test: vb/vb_parse_enum_ignore_case/test_vb_enum_comparison_operators
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

Enum Level
    Low = 10
    High = 20
End Enum

Module Program
    Sub Main()
        Dim l1 = Level.Low
        Dim l2 = Level.High
        __Check(CStr((l1 < l2) & "|" & (l1 = Level.Low)), "True|True")
    End Sub
End Module
