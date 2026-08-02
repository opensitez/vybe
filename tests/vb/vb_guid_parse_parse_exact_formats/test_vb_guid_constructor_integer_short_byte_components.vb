' vybe-test: vb/vb_guid_parse_parse_exact_formats/test_vb_guid_constructor_integer_short_byte_components
' origin: languages/vb/tests/vb/test_vb_guid_parse_parse_exact_formats.rs

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

Module Program
    Sub Main()
        ' Guid(a As Integer, b As Short, c As Short, d As Byte, e As Byte, f As Byte, g As Byte, h As Byte, i As Byte, j As Byte, k As Byte)
        Dim g As New Guid(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11)
        __Check(CStr(g.ToString("N").StartsWith("00000001")), "True")
    End Sub
End Module
