' vybe-test: vb/vb_system_guid_matrix/guid_parse_and_to_string_roundtrip
' origin: languages/vb/tests/vb/test_vb_system_guid_matrix.rs

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
        Dim parsed As Guid = Guid.Parse("4f7f1dcb-7a39-4f9b-b0de-f7f9a2f5f8f4")
        Dim text As String = parsed.ToString("N")
        __Check(CStr(text.Length), "32")
        __Check(CStr(Guid.Parse(text).ToString("D")), "4f7f1dcb-7a39-4f9b-b0de-f7f9a2f5f8f4")
    End Sub
End Module
