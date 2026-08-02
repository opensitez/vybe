' vybe-test: vb/vb_system_guid_matrix/guid_from_byte_array_roundtrip
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
        Dim source As Byte() = {
            &H4F, &H7F, &H1D, &HCB, &H7A, &H39, &H4F, &H9B,
            &HB0, &HDE, &HF7, &HF9, &HA2, &HF5, &HF8, &HF4
        }
        Dim g As New Guid(source)
        Dim bytes() As Byte = g.ToByteArray()
        Dim restored As New Guid(bytes)
        __Check(CStr(g = restored), "True")
        __Check(CStr(bytes.Length), "16")
    End Sub
End Module
