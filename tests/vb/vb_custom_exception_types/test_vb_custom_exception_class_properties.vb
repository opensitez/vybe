' vybe-test: vb/vb_custom_exception_types/test_vb_custom_exception_class_properties
' origin: languages/vb/tests/vb/test_vb_custom_exception_types.rs

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

Class ValidationException
    Inherits Exception
    Public ReadOnly FieldName As String

    Public Sub New(fieldName As String, message As String)
        MyBase.New(message)
        Me.FieldName = fieldName
    End Sub
End Class

Module Program
    Sub Main()
        Try
            Throw New ValidationException("Email", "Invalid email address format")
        Catch ex As ValidationException
            __Check(CStr(ex.FieldName & ":" & ex.Message), "Email:Invalid email address format")
        End Try
    End Sub
End Module
