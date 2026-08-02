' vybe-test: vb/vb_observable_collection_reset/test_vb_observable_collection_unsubscribing_collection_changed
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
Imports System.Collections.Specialized

Module Program
    Sub Main()
        Dim col As New ObservableCollection(Of String)()
        Dim count = 0
        Dim handler As NotifyCollectionChangedEventHandler = Sub(s, e) count += 1

        AddHandler col.CollectionChanged, handler
        col.Add("A")
        RemoveHandler col.CollectionChanged, handler
        col.Add("B")
        __Check(CStr(count), "1")
    End Sub
End Module
