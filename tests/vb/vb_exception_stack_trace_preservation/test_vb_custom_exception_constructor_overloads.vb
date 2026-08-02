' vybe-test: vb/vb_exception_stack_trace_preservation/test_vb_custom_exception_constructor_overloads
' origin: languages/vb/tests/vb/test_vb_exception_stack_trace_preservation.rs

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

Class CustomException
    Inherits Exception
    Public Sub New() : MyBase.New("Default Message") : End Sub
    Public Sub New(msg As String) : MyBase.New(msg) : End Sub
    Public Sub New(msg As String, inner As Exception) : MyBase.New(msg, inner) : End Sub
End Class

Module Program
    Sub Main()
        Dim e1 As New CustomException()
        Dim e2 As New CustomException("Custom Msg")
        Dim e3 As New CustomException("Wrapper", e1)
        __Check(CStr(e1.Message & "|" & e2.Message & "|" & e3.InnerException.Message), "Default Message|Custom Msg|Default Message")
    End Sub
End Module
