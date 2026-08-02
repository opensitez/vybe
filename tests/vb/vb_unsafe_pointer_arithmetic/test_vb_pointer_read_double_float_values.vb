' vybe-test: vb/vb_unsafe_pointer_arithmetic/test_vb_pointer_read_double_float_values
' origin: languages/vb/tests/vb/test_vb_unsafe_pointer_arithmetic.rs

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
Imports System.Runtime.InteropServices

Module Program
    Sub Main()
        Dim ptr As IntPtr = Marshal.AllocHGlobal(8)
        Dim d As Double = 99.87654
        Dim bits = BitConverter.DoubleToInt64Bits(d)
        Marshal.WriteInt64(ptr, bits)

        Dim readBits = Marshal.ReadInt64(ptr)
        Dim restoredD = BitConverter.Int64BitsToDouble(readBits)
        Marshal.FreeHGlobal(ptr)
        __Check(CStr(restoredD), "99.87654")
    End Sub
End Module
