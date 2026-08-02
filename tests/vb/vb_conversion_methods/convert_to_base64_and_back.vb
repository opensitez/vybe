' vybe-test: vb/vb_conversion_methods/convert_to_base64_and_back
' origin: languages/vb/tests/vb/test_vb_conversion_methods.rs

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

Module M
    Sub Main()
        Dim text As String = "VB.NET"
        Dim bytes As Byte() = Encoding.UTF8.GetBytes(text)
        Dim encoded As String = Convert.ToBase64String(bytes)
        Dim decoded As String = Encoding.UTF8.GetString(Convert.FromBase64String(encoded))
        __Check(CStr(encoded), "VkIuTkVUDQ==")
        __Check(CStr(decoded), "VB.NET")
    End Sub
End Module
