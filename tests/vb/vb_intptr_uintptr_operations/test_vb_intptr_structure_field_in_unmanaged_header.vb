' vybe-test: vb/vb_intptr_uintptr_operations/test_vb_intptr_structure_field_in_unmanaged_header
' origin: languages/vb/tests/vb/test_vb_intptr_uintptr_operations.rs

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

<StructLayout(LayoutKind.Sequential)>
Structure NativeHeader
    Public Length As Integer
    Public DataPtr As IntPtr
End Structure

Module Program
    Sub Main()
        Dim h As New NativeHeader With {.Length = 64, .DataPtr = New IntPtr(12345)}
        __Check(CStr(h.Length & "|" & h.DataPtr.ToInt32()), "64|12345")
    End Sub
End Module
