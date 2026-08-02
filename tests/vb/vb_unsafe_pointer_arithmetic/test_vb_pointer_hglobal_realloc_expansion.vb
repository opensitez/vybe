' vybe-test: vb/vb_unsafe_pointer_arithmetic/test_vb_pointer_hglobal_realloc_expansion
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
        Marshal.WriteInt32(ptr, 99)

        Dim newPtr As IntPtr = Marshal.ReAllocHGlobal(ptr, CType(8, IntPtr))
        Dim val1 = Marshal.ReadInt32(newPtr)
        Marshal.WriteInt32(IntPtr.Add(newPtr, 4), 100)
        Dim val2 = Marshal.ReadInt32(IntPtr.Add(newPtr, 4))
        Marshal.FreeHGlobal(newPtr)

        __Check(CStr(val1 & "|" & val2), "99|100")
    End Sub
End Module
