' vybe-test: vb/vb_directcast_value_types/test_vb_directcast_array_element_type_widening_throws
' origin: languages/vb/tests/vb/test_vb_directcast_value_types.rs

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

Module Program
    Sub Main()
        Dim bytes As Byte() = {1, 2, 3}
        Dim obj As Object = bytes
        Try
            ' Byte() cannot be DirectCast to Integer()!
            Dim ints As Integer() = DirectCast(obj, Integer())
        Catch ex As InvalidCastException
            __Check(CStr("InvalidCastException Caught on Array Type Mismatch"), "InvalidCastException Caught on Array Type Mismatch")
        End Try
    End Sub
End Module
