' vybe-test: vb/vb_bitconverter_primitive_bytes/test_vb_bitconverter_to_int64_long
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
        Dim bytes As Byte() = BitConverter.GetBytes(-9000000000000000000L)
        Dim restored = BitConverter.ToInt64(bytes, 0)
        __Check(CStr(restored), "-9000000000000000000")
    End Sub
End Module
