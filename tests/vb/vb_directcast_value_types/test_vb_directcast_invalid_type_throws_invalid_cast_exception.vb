' vybe-test: vb/vb_directcast_value_types/test_vb_directcast_invalid_type_throws_invalid_cast_exception
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

Class A
End Class

Class B
End Class

Module Program
    Sub Main()
        Dim obj As Object = New A()
        Try
            Dim b As B = DirectCast(obj, B)
        Catch ex As InvalidCastException
            __Check(CStr("InvalidCastException Caught on DirectCast"), "InvalidCastException Caught on DirectCast")
        End Try
    End Sub
End Module
