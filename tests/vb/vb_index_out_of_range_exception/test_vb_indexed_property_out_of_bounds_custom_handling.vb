' vybe-test: vb/vb_index_out_of_range_exception/test_vb_indexed_property_out_of_bounds_custom_handling
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

Class SafeArray
    Private data(2) As Integer
    Default Public Property Item(idx As Integer) As Integer
        Get
            If idx < 0 OrElse idx >= data.Length Then
                Throw New IndexOutOfRangeException("SafeArray index out of bounds")
            End If
            Return data(idx)
        End Get
        Set(value As Integer)
            If idx < 0 OrElse idx >= data.Length Then
                Throw New IndexOutOfRangeException("SafeArray index out of bounds")
            End If
            data(idx) = value
        End Set
    End Property
End Class

Module Program
    Sub Main()
        Dim sa As New SafeArray()
        Try
            sa(10) = 42
        Catch ex As IndexOutOfRangeException
            __Check(CStr(ex.Message), "SafeArray index out of bounds")
        End Try
    End Sub
End Module
