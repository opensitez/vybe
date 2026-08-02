' vybe-test: vb/vb_convert_to_base64_string/test_vb_convert_to_base64_string_offset_and_length
' origin: languages/vb/tests/vb/test_vb_convert_to_base64_string.rs

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
        Dim raw As Byte() = {10, 20, 30, 40, 50}
        ' Convert slice starting at index 1 for length 3
        Dim b64 = Convert.ToBase64String(raw, 1, 3)
        Dim restored As Byte() = Convert.FromBase64String(b64)
        __Check(CStr(String.Join(",", restored)), "20,30,40")
    End Sub
End Module
