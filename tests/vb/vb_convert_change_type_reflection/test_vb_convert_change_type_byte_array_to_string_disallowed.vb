' vybe-test: vb/vb_convert_change_type_reflection/test_vb_convert_change_type_byte_array_to_string_disallowed
' origin: languages/vb/tests/vb/test_vb_convert_change_type_reflection.rs

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
        Dim bytes As Object = New Byte() {65, 66}
        Try
            ' Direct ChangeType from Byte() to String throws InvalidCastException!
            Convert.ChangeType(bytes, GetType(String))
        Catch ex As InvalidCastException
            __Check(CStr("InvalidCastException Caught on Byte Array to String"), "InvalidCastException Caught on Byte Array to String")
        End Try
    End Sub
End Module
