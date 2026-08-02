' vybe-test: vb/vb_system_conversion_builtins_matrix/conversion_builtins_base64_roundtrip
' origin: languages/vb/tests/vb/test_vb_system_conversion_builtins_matrix.rs

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
        Dim payload() As Byte = Encoding.UTF8.GetBytes("vb")
        Dim encoded As String = Convert.ToBase64String(payload)
        Dim decoded() As Byte = Convert.FromBase64String(encoded)

        __Check(CStr(encoded.Length > 0), "True")
        __Check(CStr(Encoding.UTF8.GetString(decoded)), "vb")
        __Check(CStr(Convert.ToBase64String(Encoding.UTF8.GetBytes("")) = ""), "True")
    End Sub
End Module
