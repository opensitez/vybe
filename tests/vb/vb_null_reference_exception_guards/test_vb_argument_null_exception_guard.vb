' vybe-test: vb/vb_null_reference_exception_guards/test_vb_argument_null_exception_guard
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

Module Program
    Private Sub ProcessData(data As String)
        If data Is Nothing Then
            Throw New ArgumentNullException(NameOf(data), "Data cannot be null")
        End If
    End Sub

    Sub Main()
        Try
            ProcessData(Nothing)
        Catch ex As ArgumentNullException
            __Check(CStr("Param: " & ex.ParamName), "Param: data")
        End Try
    End Sub
End Module
