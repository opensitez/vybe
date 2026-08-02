' vybe-test: vb/vb_unsafe_pointer_arithmetic/test_vb_pointer_offset_calculation_int32_array
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
        Dim numbers As Integer() = {100, 200, 300, 400}
        Dim handle = GCHandle.Alloc(numbers, GCHandleType.Pinned)
        Dim baseAddr = handle.AddrOfPinnedObject()

        ' Offset for element index 2: 2 * Marshal.SizeOf(GetType(Integer)) = 8 bytes
        Dim elem2Addr = IntPtr.Add(baseAddr, 2 * 4)
        Dim val = Marshal.ReadInt32(elem2Addr)
        handle.Free()
        __Check(CStr(val), "300")
    End Sub
End Module
