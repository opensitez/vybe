' vybe-test: vb/vb_observable_collection_reset/test_vb_observable_collection_multiple_listeners
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

Imports System.Collections.ObjectModel

Module Program
    Sub Main()
        Dim col As New ObservableCollection(Of String)()
        Dim c1 = 0, c2 = 0
        AddHandler col.CollectionChanged, Sub(s, e) c1 += 1
        AddHandler col.CollectionChanged, Sub(s, e) c2 += 1

        col.Add("A")
        col.Add("B")
        __Check(CStr(c1 & "|" & c2), "2|2")
    End Sub
End Module
