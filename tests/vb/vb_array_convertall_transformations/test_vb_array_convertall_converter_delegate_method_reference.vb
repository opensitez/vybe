' vybe-test: vb/vb_array_convertall_transformations/test_vb_array_convertall_converter_delegate_method_reference
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

Module Converter
    Public Function DoubleVal(n As Integer) As Integer
        Return n * 2
    End Function
End Module

Module Program
    Sub Main()
        Dim numbers As Integer() = {1, 2, 3}
        Dim doubled As Integer() = Array.ConvertAll(numbers, AddressOf Converter.DoubleVal)
        __Check(CStr(String.Join(",", doubled)), "2,4,6")
    End Sub
End Module
