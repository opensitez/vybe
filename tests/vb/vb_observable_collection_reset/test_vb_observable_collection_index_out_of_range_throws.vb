' vybe-test: vb/vb_observable_collection_reset/test_vb_observable_collection_index_out_of_range_throws
' origin: languages/vb/tests/vb/test_vb_observable_collection_reset.rs

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
Imports System.Collections.ObjectModel

Module Program
    Sub Main()
        Dim col As New ObservableCollection(Of Integer) From {1, 2}
        Try
            col.RemoveAt(5)
        Catch ex As ArgumentOutOfRangeException
            __Check(CStr("ArgumentOutOfRangeException Caught"), "ArgumentOutOfRangeException Caught")
        End Try
    End Sub
End Module
