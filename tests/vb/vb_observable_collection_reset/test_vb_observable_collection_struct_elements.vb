' vybe-test: vb/vb_observable_collection_reset/test_vb_observable_collection_struct_elements
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

Structure Point2D
    Public X, Y As Integer
End Structure

Module Program
    Sub Main()
        Dim col As New ObservableCollection(Of Point2D)()
        Dim addedPt As Point2D
        AddHandler col.CollectionChanged, Sub(s, e)
            If e.NewItems IsNot Nothing Then addedPt = CType(e.NewItems(0), Point2D)
        End Sub

        col.Add(New Point2D With {.X = 5, .Y = 10})
        __Check(CStr(addedPt.X & "," & addedPt.Y), "5,10")
    End Sub
End Module
