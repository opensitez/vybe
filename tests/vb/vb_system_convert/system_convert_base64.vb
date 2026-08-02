' vybe-test: vb/vb_system_convert/system_convert_base64
' origin: languages/vb/tests/vb/test_vb_system_convert.rs

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
        Dim original As String = "VB.NET Rocks!"
        Dim bytes As Byte() = Encoding.UTF8.GetBytes(original)
        
        Dim base64 As String = Convert.ToBase64String(bytes)
        __Check(CStr(base64 IsNot Nothing), "True")
        
        Dim decodedBytes As Byte() = Convert.FromBase64String(base64)
        Dim decoded As String = Encoding.UTF8.GetString(decodedBytes)
        
        __Check(CStr(decoded = original), "True")
    End Sub
End Module
