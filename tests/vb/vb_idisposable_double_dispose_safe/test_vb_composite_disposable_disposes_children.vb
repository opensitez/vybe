' vybe-test: vb/vb_idisposable_double_dispose_safe/test_vb_composite_disposable_disposes_children
' origin: languages/vb/tests/vb/test_vb_idisposable_double_dispose_safe.rs

Imports System
Imports System.Collections.Generic

Class CompositeDisposable
    Implements IDisposable
    Private children As New List(Of IDisposable)()

    Public Sub Add(item As IDisposable)
        children.Add(item)
    End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        For Each child In children
            child.Dispose()
        Next
        children.Clear()
    End Sub
End Class

Class ChildRes
    Implements IDisposable
    Private tag As String
    Public Sub New(t As String)
        tag = t
    End Sub
    Public Sub Dispose() Implements IDisposable.Dispose
        Console.WriteLine("Disposed Child " & tag)
    End Sub
End Class

Module Program
    Sub Main()
        Dim comp As New CompositeDisposable()
        comp.Add(New ChildRes("A"))
        comp.Add(New ChildRes("B"))
        comp.Dispose()
    End Sub
End Module
