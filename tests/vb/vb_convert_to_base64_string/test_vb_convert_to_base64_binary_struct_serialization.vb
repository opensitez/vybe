' vybe-test: vb/vb_convert_to_base64_string/test_vb_convert_to_base64_binary_struct_serialization
' origin: languages/vb/tests/vb/test_vb_convert_to_base64_string.rs

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
Structure RecordHeader
    Public Magic As Integer
    Public Length As Short
End Structure

Module Program
    Sub Main()
        Dim h As New RecordHeader With {.Magic = &H41424344, .Length = 100}
        Dim size = Marshal.SizeOf(GetType(RecordHeader))
        Dim ptr = Marshal.AllocHGlobal(size)
        Marshal.StructureToPtr(h, ptr, False)

        Dim bytes(size - 1) As Byte
        Marshal.Copy(ptr, bytes, 0, size)
        Marshal.FreeHGlobal(ptr)

        Dim b64 = Convert.ToBase64String(bytes)
        __Check(CStr(b64.Length > 0), "True")
    End Sub
End Module
