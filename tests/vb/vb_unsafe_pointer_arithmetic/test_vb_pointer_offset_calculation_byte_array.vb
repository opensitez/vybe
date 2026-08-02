' vybe-test: vb/vb_unsafe_pointer_arithmetic/test_vb_pointer_offset_calculation_byte_array
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
        Dim bytes As Byte() = {10, 20, 30, 40, 50}
        Dim handle = GCHandle.Alloc(bytes, GCHandleType.Pinned)
        Dim baseAddr = handle.AddrOfPinnedObject()

        Dim thirdByteAddr = IntPtr.Add(baseAddr, 2)
        Dim val = Marshal.ReadByte(thirdByteAddr)
        handle.Free()
        __Check(CStr(val), "30")
    End Sub
End Module
