' vybe-test: vb/vb_casting_patterns/string_and_char_casting_via_ctype
' origin: languages/vb/tests/vb/test_vb_casting_patterns.rs

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

Module M
    Sub Main()
        Dim o As Object = "x"
        Dim s As String = CType(o, String)
        Dim c As Char = CType(s(0), Char)
        __Check(CStr(s), "x")
        __Check(CStr(c), "x")
    End Sub
End Module
