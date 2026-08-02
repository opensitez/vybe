' vybe-test: vb/vb_convert_to_base64_string/test_vb_convert_to_base64_url_safe_replacement_simulation
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
    Private Function ToBase64Url(data As Byte()) As String
        Dim base64 = Convert.ToBase64String(data)
        Return base64.Replace("+", "-").Replace("/", "_").TrimEnd("="c)
    End Function

    Sub Main()
        Dim bytes As Byte() = Encoding.UTF8.GetBytes("Subject?Data#1")
        Dim urlSafe = ToBase64Url(bytes)
        __Check(CStr(Not urlSafe.Contains("+") AndAlso Not urlSafe.Contains("/")), "True")
    End Sub
End Module
