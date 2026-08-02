' vybe-test: vb/vb_types/enum_basic
' origin: languages/vb/tests/vb/test_vb_types.rs

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

Enum Color
    Red
    Green
    Blue
End Enum

Module M
    Sub Main()
        Dim c As Color = Color.Green
        __Check(CStr(c), "1")
    End Sub
End Module
