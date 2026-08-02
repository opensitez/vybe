' vybe-test: vb/vb_safe_handle_invalid_check/test_vb_safe_handle_marshal_structure_to_ptr_with_handle
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

<StructLayout(LayoutKind.Sequential)>
Structure NativeConfig
    Public Version As Integer
    Public HandleVal As IntPtr
End Structure

Module Program
    Sub Main()
        Dim cfg As New NativeConfig With {.Version = 1, .HandleVal = New IntPtr(99)}
        Dim size = Marshal.SizeOf(GetType(NativeConfig))
        Dim mem = Marshal.AllocHGlobal(size)
        Marshal.StructureToPtr(cfg, mem, False)

        Dim readBack As NativeConfig = CType(Marshal.PtrToStructure(mem, GetType(NativeConfig)), NativeConfig)
        Marshal.FreeHGlobal(mem)

        __Check(CStr(readBack.Version & "|" & readBack.HandleVal.ToInt64()), "1|99")
    End Sub
End Module
