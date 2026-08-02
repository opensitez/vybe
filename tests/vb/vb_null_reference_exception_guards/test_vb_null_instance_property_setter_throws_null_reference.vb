' vybe-test: vb/vb_null_reference_exception_guards/test_vb_null_instance_property_setter_throws_null_reference
' origin: languages/vb/tests/vb/test_vb_null_reference_exception_guards.rs

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

Class User
    Public Property Name As String
End Class

Module Program
    Sub Main()
        Dim u As User = Nothing
        Try
            u.Name = "Alice"
        Catch ex As NullReferenceException
            __Check(CStr("NullReferenceException Caught on Property Set"), "NullReferenceException Caught on Property Set")
        End Try
    End Sub
End Module
