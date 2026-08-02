' vybe-test: vb/vb_unsafe_pointer_arithmetic/test_vb_pointer_offset_write_int32
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
        Dim numbers As Integer() = {0, 0, 0}
        Dim handle = GCHandle.Alloc(numbers, GCHandleType.Pinned)
        Dim baseAddr = handle.AddrOfPinnedObject()

        Marshal.WriteInt32(IntPtr.Add(baseAddr, 0), 10)
        Marshal.WriteInt32(IntPtr.Add(baseAddr, 4), 20)
        Marshal.WriteInt32(IntPtr.Add(baseAddr, 8), 30)
        handle.Free()
        __Check(CStr(String.Join(",", numbers)), "10,20,30")
    End Sub
End Module
