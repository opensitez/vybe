' vybe-test: vb/vb_observable_collection_reset/test_vb_observable_collection_custom_item_type
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

Class TaskItem
    Public Title As String
End Class

Module Program
    Sub Main()
        Dim col As New ObservableCollection(Of TaskItem)()
        Dim addedTitle = ""
        AddHandler col.CollectionChanged, Sub(s, e)
            If e.NewItems IsNot Nothing Then
                addedTitle = CType(e.NewItems(0), TaskItem).Title
            End If
        End Sub

        col.Add(New TaskItem With {.Title = "BuildApp"})
        __Check(CStr(addedTitle), "BuildApp")
    End Sub
End Module
