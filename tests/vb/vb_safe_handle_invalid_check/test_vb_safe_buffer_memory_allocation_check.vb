' vybe-test: vb/vb_safe_handle_invalid_check/test_vb_safe_buffer_memory_allocation_check
' origin: languages/vb/tests/vb/test_vb_safe_handle_invalid_check.rs

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
        ' SafeBuffer interop check
        Dim ptr = Marshal.AllocHGlobal(128)
        Marshal.FreeHGlobal(ptr)
        __Check(CStr("AllocHGlobal Safe"), "AllocHGlobal Safe")
    End Sub
End Module
