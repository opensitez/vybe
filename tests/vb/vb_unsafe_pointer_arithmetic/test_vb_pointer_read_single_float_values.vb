' vybe-test: vb/vb_unsafe_pointer_arithmetic/test_vb_pointer_read_single_float_values
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
        Dim ptr As IntPtr = Marshal.AllocHGlobal(4)
        Dim f As Single = 12.34F
        Dim bits = BitConverter.SingleToInt32Bits(f)
        Marshal.WriteInt32(ptr, bits)

        Dim readBits = Marshal.ReadInt32(ptr)
        Dim restoredF = BitConverter.Int32BitsToSingle(readBits)
        Marshal.FreeHGlobal(ptr)
        __Check(CStr(restoredF), "12.34")
    End Sub
End Module
