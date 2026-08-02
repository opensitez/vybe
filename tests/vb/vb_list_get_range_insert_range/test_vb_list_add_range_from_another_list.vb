' vybe-test: vb/vb_list_get_range_insert_range/test_vb_list_add_range_from_another_list
' origin: languages/vb/tests/vb/test_vb_list_get_range_insert_range.rs

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

Module Program
    Sub Main()
        Dim srcList As New List(Of Double) From {1.1, 2.2}
        Dim destList As New List(Of Double) From {0.0}
        destList.AddRange(srcList)
        __Check(CStr(String.Join(";", destList)), "0;1.1;2.2")
    End Sub
End Module
