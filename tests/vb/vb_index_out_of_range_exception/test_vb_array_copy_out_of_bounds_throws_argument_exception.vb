' vybe-test: vb/vb_index_out_of_range_exception/test_vb_array_copy_out_of_bounds_throws_argument_exception
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
        Dim srcArr As Integer() = {1, 2, 3, 4, 5}
        Dim destArr As Integer() = new Integer(2) {}
        Try
            ' Trying to copy 5 elements into destination of size 3
            Array.Copy(srcArr, destArr, 5)
        Catch ex As ArgumentException
            __Check(CStr("Array.Copy ArgumentException Caught"), "Array.Copy ArgumentException Caught")
        End Try
    End Sub
End Module
