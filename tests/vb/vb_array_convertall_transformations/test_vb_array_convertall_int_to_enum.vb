' vybe-test: vb/vb_array_convertall_transformations/test_vb_array_convertall_int_to_enum
' origin: languages/vb/tests/vb/test_vb_array_convertall_transformations.rs

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

Enum Mode
    Off = 0
    OnVal = 1
End Enum

Module Program
    Sub Main()
        Dim raw As Integer() = {0, 1, 0}
        Dim modes As Mode() = Array.ConvertAll(raw, Function(r) CType(r, Mode))
        __Check(CStr(modes(0).ToString() & "," & modes(1).ToString()), "Off,OnVal")
    End Sub
End Module
