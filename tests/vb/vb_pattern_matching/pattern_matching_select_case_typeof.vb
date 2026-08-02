' vybe-test: vb/vb_pattern_matching/pattern_matching_select_case_typeof
' origin: languages/vb/tests/vb/test_vb_pattern_matching.rs

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

Module M
    Sub PrintType(obj As Object)
        Select Case obj
            Case i As Integer
                __Check(CStr("Integer: " & i.ToString()), "Integer: 42")
            Case s As String
                __Check(CStr("String: " & s), "String: Hello")
            Case Else
                __Check(CStr("Unknown"), "Unknown")
        End Select
    End Sub

    Sub Main()
        PrintType(42)
        PrintType("Hello")
        PrintType(5.5)
    End Sub
End Module
