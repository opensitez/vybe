' vybe-test: vb/vb_directcast_value_types/test_vb_directcast_boxed_value_type_widening_throws_invalid_cast
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
        Dim boxedInt As Object = 42
        Try
            ' DirectCast to Double from boxed Integer throws InvalidCastException (unlike CType)!
            Dim d As Double = DirectCast(boxedInt, Double)
        Catch ex As InvalidCastException
            __Check(CStr("InvalidCastException Caught on DirectCast Widening"), "InvalidCastException Caught on DirectCast Widening")
        End Try
    End Sub
End Module
