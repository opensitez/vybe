' vybe-test: vb/vb_system_exception_types_matrix/exception_argument_out_of_range_exposes_actual_value
' origin: languages/vb/tests/vb/test_vb_system_exception_types_matrix.rs

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

Module M
    Sub Main()
        Try
            Throw New ArgumentOutOfRangeException("index", 13, "value must be non-negative")
        Catch ex As ArgumentOutOfRangeException
            __Check(CStr(ex.ParamName), "index")
            __Check(CStr(ex.ActualValue = 13), "True")
            __Check(CStr(ex.GetType().Name), "ArgumentOutOfRangeException")
        End Try
    End Sub
End Module
