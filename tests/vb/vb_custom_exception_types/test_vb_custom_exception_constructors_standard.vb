' vybe-test: vb/vb_custom_exception_types/test_vb_custom_exception_constructors_standard
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

Class CustomNotFoundException
    Inherits Exception

    Public Sub New()
        MyBase.New("Resource not found")
    End Sub

    Public Sub New(message As String)
        MyBase.New(message)
    End Sub

    Public Sub New(message As String, inner As Exception)
        MyBase.New(message, inner)
    End Sub
End Class

Module Program
    Sub Main()
        Try
            Try
                Throw New InvalidOperationException("Inner error")
            Catch innerEx As Exception
                Throw New CustomNotFoundException("Outer error", innerEx)
            End Try
        Catch outerEx As CustomNotFoundException
            __Check(CStr(outerEx.Message), "Outer error")
            __Check(CStr(outerEx.InnerException.Message), "Inner error")
        End Try
    End Sub
End Module
