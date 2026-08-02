' vybe-test: vb/vb_system_encoding/ascii_roundtrip_for_plain_ascii
' origin: languages/vb/tests/vb/test_vb_system_encoding.rs

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

Imports System.Text

Module M
    Sub Main()
        Dim bytes As Byte() = Encoding.ASCII.GetBytes("Hello")
        __Check(CStr(Encoding.ASCII.GetString(bytes)), "Hello")
        __Check(CStr(bytes.Length), "5")
    End Sub
End Module
