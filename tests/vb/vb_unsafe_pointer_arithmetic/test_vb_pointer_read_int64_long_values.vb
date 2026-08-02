' vybe-test: vb/vb_unsafe_pointer_arithmetic/test_vb_pointer_read_int64_long_values
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
        Marshal.WriteInt64(ptr, 0, 5000000000L)
        Dim l1 = Marshal.ReadInt64(ptr, 0)
        Marshal.FreeHGlobal(ptr)
        __Check(CStr(l1), "5000000000")
    End Sub
End Module
