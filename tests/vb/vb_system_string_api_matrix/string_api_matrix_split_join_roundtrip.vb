' vybe-test: vb/vb_system_string_api_matrix/string_api_matrix_split_join_roundtrip
' origin: languages/vb/tests/vb/test_vb_system_string_api_matrix.rs

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
        Dim text As String = "a,b;c,d"
        Dim pieces As String() = text.Split(","c, ";"c)

        Dim joined As String = String.Join("|", pieces)
        __Check(CStr(pieces.Length), "4")
        __Check(CStr(joined), "a|b|c|d")
        __Check(CStr(pieces(2)), "c")
        __Check(CStr(pieces(3)), "d")
    End Sub
End Module
