' vybe-test: vb/vb_convert_to_base64_string/test_vb_convert_from_base64_url_safe_restoration
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
Imports System.Text

Module Program
    Private Function FromBase64Url(base64Url As String) As Byte()
        Dim padded = base64Url.Replace("-", "+").Replace("_", "/")
        Select Case padded.Length Mod 4
            Case 2 : padded &= "=="
            Case 3 : padded &= "="
        End Select
        Return Convert.FromBase64String(padded)
    End Function

    Sub Main()
        Dim bytes = FromBase64Url("U3ViamVjdD9EYXRhIzE")
        __Check(CStr(Encoding.UTF8.GetString(bytes)), "Subject?Data#1")
    End Sub
End Module
