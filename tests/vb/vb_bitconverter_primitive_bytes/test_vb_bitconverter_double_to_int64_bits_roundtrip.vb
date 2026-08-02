' vybe-test: vb/vb_bitconverter_primitive_bytes/test_vb_bitconverter_double_to_int64_bits_roundtrip
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
        Dim dbl = 2.718281828459
        Dim bits = BitConverter.DoubleToInt64Bits(dbl)
        Dim restored = BitConverter.Int64BitsToDouble(bits)
        __Check(CStr(dbl = restored), "True")
    End Sub
End Module
