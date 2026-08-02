' vybe-test: vb/vb_list_get_range_insert_range/test_vb_list_get_range_independence_from_original
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
        Dim original As New List(Of Integer) From {1, 2, 3}
        Dim subList As List(Of Integer) = original.GetRange(0, 2)
        subList(0) = 99
        __Check(CStr(original(0) & ":" & subList(0)), "1:99")
    End Sub
End Module
