' vybe-test: vb/vb_bitconverter_primitive_bytes/test_vb_bitconverter_get_bytes_boolean
' origin: languages/vb/tests/vb/test_vb_bitconverter_primitive_bytes.rs

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
        Dim bytesTrue = BitConverter.GetBytes(True)
        Dim bytesFalse = BitConverter.GetBytes(False)
        __Check(CStr(bytesTrue(0) & "|" & bytesFalse(0)), "1|0")
    End Sub
End Module
