' vybe-test: vb/vb_struct_layout_sequential/test_vb_struct_layout_char_set_unicode
' origin: languages/vb/tests/vb/test_vb_struct_layout_sequential.rs

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

Imports System.Runtime.InteropServices

<StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)>
Structure UnicodeStringStruct
    <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=10)>
    Public Text As String
End Structure

Module Program
    Sub Main()
        ' 10 WChars * 2 bytes = 20 bytes
        __Check(CStr(Marshal.SizeOf(GetType(UnicodeStringStruct))), "20")
    End Sub
End Module
