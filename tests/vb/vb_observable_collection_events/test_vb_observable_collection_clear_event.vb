' vybe-test: vb/vb_observable_collection_events/test_vb_observable_collection_clear_event
' origin: languages/vb/tests/vb/test_vb_observable_collection_events.rs

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
        Dim collection As New ObservableCollection(Of Integer) From {1, 2, 3}
        AddHandler collection.CollectionChanged, Sub(sender, e)
            __Check(CStr("Action: " & e.Action.ToString()), "Action: Reset")
        End Sub

        collection.Clear()
    End Sub
End Module
