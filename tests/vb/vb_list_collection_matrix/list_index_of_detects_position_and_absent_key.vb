' vybe-test: vb/vb_list_collection_matrix/list_index_of_detects_position_and_absent_key
' origin: languages/vb/tests/vb/test_vb_list_collection_matrix.rs

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

Imports System.Collections.Generic

Module M
    Sub Main()
        Dim values As New List(Of Integer) From {7, 8, 9}
        __Check(CStr(values.IndexOf(8)), "1")
        __Check(CStr(values.IndexOf(99)), "-1")
    End Sub
End Module
