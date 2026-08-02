' vybe-test: vb/vb_system_bitconverter_matrix/bitconverter_roundtrip_double_bytes
' origin: languages/vb/tests/vb/test_vb_system_bitconverter_matrix.rs

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
        Dim source As Double = 12.5
        Dim bytes() As Byte = BitConverter.GetBytes(source)
        __Check(CStr(BitConverter.ToDouble(bytes, 0)), "12.5")
        __Check(CStr(bytes.Length), "8")
    End Sub
End Module
