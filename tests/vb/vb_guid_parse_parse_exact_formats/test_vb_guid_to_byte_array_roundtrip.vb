' vybe-test: vb/vb_guid_parse_parse_exact_formats/test_vb_guid_to_byte_array_roundtrip
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
        Dim orig = Guid.NewGuid()
        Dim bytes = orig.ToByteArray()
        Dim restored As New Guid(bytes)
        __Check(CStr((orig = restored) & "|" & (bytes.Length = 16)), "True|True")
    End Sub
End Module
