' vybe-test: vb/vb_marshal_size_of_structure/test_vb_marshal_offset_of_struct_field
' origin: languages/vb/tests/vb/test_vb_marshal_size_of_structure.rs

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
Structure CompoundStruct
    Public A As Byte
    Public B As Integer
End Structure

Module Program
    Sub Main()
        ' Alignment padding: Offset of B should be 4 bytes due to 4-byte int alignment!
        Dim offsetB = Marshal.OffsetOf(GetType(CompoundStruct), "B")
        __Check(CStr(offsetB.ToInt32()), "4")
    End Sub
End Module
