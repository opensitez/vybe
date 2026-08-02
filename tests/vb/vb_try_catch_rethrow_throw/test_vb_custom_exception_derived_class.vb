' vybe-test: vb/vb_try_catch_rethrow_throw/test_vb_custom_exception_derived_class
' origin: languages/vb/tests/vb/test_vb_try_catch_rethrow_throw.rs

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

Class BusinessRuleException
    Inherits Exception
    Public Property ErrorCode As Integer
    Public Sub New(code As Integer, msg As String)
        MyBase.New(msg)
        ErrorCode = code
    End Sub
End Class

Module Program
    Sub Main()
        Try
            Throw New BusinessRuleException(404, "Entity Not Found")
        Catch ex As BusinessRuleException
            __Check(CStr("Code: " & ex.ErrorCode & " | Msg: " & ex.Message), "Code: 404 | Msg: Entity Not Found")
        End Try
    End Sub
End Module
