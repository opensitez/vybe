' vybe-test: vb/vb_index_out_of_range_exception/test_vb_multidimensional_array_dimension_0_out_of_bounds
' origin: languages/vb/tests/vb/test_vb_index_out_of_range_exception.rs

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
        Dim mat(1, 1) As Integer
        Try
            mat(2, 0) = 5
        Catch ex As IndexOutOfRangeException
            __Check(CStr("Dim0 Out Of Bounds Caught"), "Dim0 Out Of Bounds Caught")
        End Try
    End Sub
End Module
