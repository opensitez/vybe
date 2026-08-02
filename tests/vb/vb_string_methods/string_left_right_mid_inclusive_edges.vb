' vybe-test: vb/vb_string_methods/string_left_right_mid_inclusive_edges
' origin: languages/vb/tests/vb/test_vb_string_methods.rs

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

Imports Microsoft.VisualBasic

Module M
    Sub Main()
        Dim source As String = "abcdef"
        __Check(CStr(Strings.Left(source, 2)), "ab")
        __Check(CStr(Strings.Right(source, 2)), "ef")
        __Check(CStr(Strings.Mid(source, 2, 3)), "bcd")
    End Sub
End Module
