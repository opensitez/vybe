' vybe-test: vb/vb_gchandle_alloc_free/test_vb_gchandle_alloc_pinned_byte_array
' origin: languages/vb/tests/vb/test_vb_gchandle_alloc_free.rs

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
        Dim bytes As Byte() = {10, 20, 30, 40}
        Dim handle As GCHandle = GCHandle.Alloc(bytes, GCHandleType.Pinned)
        Dim addr As IntPtr = handle.AddrOfPinnedObject()
        __Check(CStr((addr <> IntPtr.Zero) & "|" & Marshal.ReadByte(addr)), "True|10")
        handle.Free()
    End Sub
End Module
