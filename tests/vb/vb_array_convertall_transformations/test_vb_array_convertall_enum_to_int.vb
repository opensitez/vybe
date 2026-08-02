' vybe-test: vb/vb_array_convertall_transformations/test_vb_array_convertall_enum_to_int
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

Enum Level
    Low = 1
    Medium = 2
    High = 3
End Enum

Module Program
    Sub Main()
        Dim levels As Level() = {Level.Low, Level.High}
        Dim ints As Integer() = Array.ConvertAll(levels, Function(l) CInt(l))
        __Check(CStr(String.Join(",", ints)), "1,3")
    End Sub
End Module
