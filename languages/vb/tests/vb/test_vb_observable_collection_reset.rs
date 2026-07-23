use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: ObservableCollection CollectionChanged & Reset Actions
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_observable_collection_add_item_event() {
    let src = r#"
Imports System.Collections.ObjectModel
Imports System.Collections.Specialized

Module Program
    Sub Main()
        Dim col As New ObservableCollection(Of String)()
        Dim action = ""
        AddHandler col.CollectionChanged, Sub(s, e) action = e.Action.ToString()
        col.Add("Item1")
        Console.WriteLine(action)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Add"]);
}

#[test]
fn test_vb_observable_collection_remove_item_event() {
    let src = r#"
Imports System.Collections.ObjectModel
Imports System.Collections.Specialized

Module Program
    Sub Main()
        Dim col As New ObservableCollection(Of String) From {"Item1", "Item2"}
        Dim removedItem = ""
        AddHandler col.CollectionChanged, Sub(s, e)
            If e.Action = NotifyCollectionChangedAction.Remove Then
                removedItem = e.OldItems(0).ToString()
            End If
        End Sub
        col.Remove("Item1")
        Console.WriteLine(removedItem)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Item1"]);
}

#[test]
fn test_vb_observable_collection_replace_item_event() {
    let src = r#"
Imports System.Collections.ObjectModel
Imports System.Collections.Specialized

Module Program
    Sub Main()
        Dim col As New ObservableCollection(Of String) From {"Old"}
        Dim changeLog = ""
        AddHandler col.CollectionChanged, Sub(s, e)
            If e.Action = NotifyCollectionChangedAction.Replace Then
                changeLog = e.OldItems(0).ToString() & "->" & e.NewItems(0).ToString()
            End If
        End Sub
        col(0) = "New"
        Console.WriteLine(changeLog)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Old->New"]);
}

#[test]
fn test_vb_observable_collection_move_item_event() {
    let src = r#"
Imports System.Collections.ObjectModel
Imports System.Collections.Specialized

Module Program
    Sub Main()
        Dim col As New ObservableCollection(Of String) From {"A", "B", "C"}
        Dim moveInfo = ""
        AddHandler col.CollectionChanged, Sub(s, e)
            If e.Action = NotifyCollectionChangedAction.Move Then
                moveInfo = e.OldStartingIndex & "->" & e.NewStartingIndex
            End If
        End Sub
        col.Move(0, 2)
        Console.WriteLine(moveInfo)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["0->2"]);
}

#[test]
fn test_vb_observable_collection_clear_reset_event() {
    let src = r#"
Imports System.Collections.ObjectModel
Imports System.Collections.Specialized

Module Program
    Sub Main()
        Dim col As New ObservableCollection(Of Integer) From {1, 2, 3}
        Dim resetFired = False
        AddHandler col.CollectionChanged, Sub(s, e)
            If e.Action = NotifyCollectionChangedAction.Reset Then resetFired = True
        End Sub
        col.Clear()
        Console.WriteLine(resetFired)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_observable_collection_insert_starting_index() {
    let src = r#"
Imports System.Collections.ObjectModel
Imports System.Collections.Specialized

Module Program
    Sub Main()
        Dim col As New ObservableCollection(Of String) From {"A", "C"}
        Dim newIdx = -1
        AddHandler col.CollectionChanged, Sub(s, e)
            If e.Action = NotifyCollectionChangedAction.Add Then newIdx = e.NewStartingIndex
        End Sub
        col.Insert(1, "B")
        Console.WriteLine(newIdx & "|" & String.Join("", col))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1|ABC"]);
}

#[test]
fn test_vb_observable_collection_property_changed_count_and_indexer() {
    let src = r#"
Imports System.Collections.ObjectModel
Imports System.ComponentModel

Module Program
    Sub Main()
        Dim col As New ObservableCollection(Of Integer)()
        Dim propChangedList As New System.Collections.Generic.List(Of String)()
        AddHandler CType(col, INotifyPropertyChanged).PropertyChanged, Sub(s, e)
            propChangedList.Add(e.PropertyName)
        End Sub
        col.Add(100)
        Console.WriteLine(String.Join(",", propChangedList))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Count,Item[]"]);
}

#[test]
fn test_vb_custom_bulk_observable_collection_range_add() {
    let src = r#"
Imports System.Collections.Generic
Imports System.Collections.ObjectModel
Imports System.Collections.Specialized

Class RangeObservableCollection(Of T)
    Inherits ObservableCollection(Of T)

    Public Sub AddRange(items As IEnumerable(Of T))
        For Each item In items
            Items.Add(item)
        Next
        OnCollectionChanged(New NotifyCollectionChangedEventArgs(NotifyCollectionChangedAction.Reset))
    End Sub
End Class

Module Program
    Sub Main()
        Dim col As New RangeObservableCollection(Of Integer)()
        Dim resetCount = 0
        AddHandler col.CollectionChanged, Sub(s, e)
            If e.Action = NotifyCollectionChangedAction.Reset Then resetCount += 1
        End Sub
        col.AddRange({10, 20, 30})
        Console.WriteLine(col.Count & "|Reset=" & resetCount)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["3|Reset=1"]);
}

#[test]
fn test_vb_observable_collection_remove_at_index() {
    let src = r#"
Imports System.Collections.ObjectModel
Imports System.Collections.Specialized

Module Program
    Sub Main()
        Dim col As New ObservableCollection(Of String) From {"X", "Y", "Z"}
        Dim oldIdx = -1
        AddHandler col.CollectionChanged, Sub(s, e)
            If e.Action = NotifyCollectionChangedAction.Remove Then oldIdx = e.OldStartingIndex
        End Sub
        col.RemoveAt(1)
        Console.WriteLine(oldIdx & "|" & String.Join("", col))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1|XZ"]);
}

#[test]
fn test_vb_observable_collection_multiple_listeners() {
    let src = r#"
Imports System.Collections.ObjectModel

Module Program
    Sub Main()
        Dim col As New ObservableCollection(Of String)()
        Dim c1 = 0, c2 = 0
        AddHandler col.CollectionChanged, Sub(s, e) c1 += 1
        AddHandler col.CollectionChanged, Sub(s, e) c2 += 1

        col.Add("A")
        col.Add("B")
        Console.WriteLine(c1 & "|" & c2)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2|2"]);
}

#[test]
fn test_vb_observable_collection_custom_item_type() {
    let src = r#"
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
        Console.WriteLine(addedTitle)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["BuildApp"]);
}

#[test]
fn test_vb_observable_collection_item_property_changed_bubbling() {
    let src = r#"
Imports System.Collections.ObjectModel
Imports System.ComponentModel

Class NotifyingItem
    Implements INotifyPropertyChanged
    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

    Private _val As String
    Public Property Val As String
        Get
            Return _val
        End Get
        Set(v As String)
            _val = v
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs("Val"))
        End Set
    End Property
End Class

Module Program
    Sub Main()
        Dim item As New NotifyingItem()
        Dim col As New ObservableCollection(Of NotifyingItem)()
        col.Add(item)

        Dim changed = False
        AddHandler item.PropertyChanged, Sub(s, e) changed = True
        item.Val = "Updated"
        Console.WriteLine(changed)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_observable_collection_reentrant_modification_throws() {
    let src = r#"
Imports System
Imports System.Collections.ObjectModel

Module Program
    Sub Main()
        Dim col As New ObservableCollection(Of Integer)()
        AddHandler col.CollectionChanged, Sub(s, e)
            ' Modifying collection during its own CollectionChanged event raises exception!
            Try
                col.Add(999)
            Catch ex As InvalidOperationException
                Console.WriteLine("InvalidOperationException Caught on Reentrant Add")
            End Try
        End Sub
        col.Add(1)
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["InvalidOperationException Caught on Reentrant Add"]
    );
}

#[test]
fn test_vb_observable_collection_read_only_wrapper() {
    let src = r#"
Imports System.Collections.ObjectModel

Module Program
    Sub Main()
        Dim col As New ObservableCollection(Of String) From {"A"}
        Dim roCol As New ReadOnlyObservableCollection(Of String)(col)

        Dim actionFired = ""
        AddHandler CType(roCol, INotifyCollectionChanged).CollectionChanged, Sub(s, e)
            actionFired = e.Action.ToString()
        End Sub

        col.Add("B")
        Console.WriteLine(roCol.Count & "|" & actionFired)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2|Add"]);
}

#[test]
fn test_vb_observable_collection_index_out_of_range_throws() {
    let src = r#"
Imports System
Imports System.Collections.ObjectModel

Module Program
    Sub Main()
        Dim col As New ObservableCollection(Of Integer) From {1, 2}
        Try
            col.RemoveAt(5)
        Catch ex As ArgumentOutOfRangeException
            Console.WriteLine("ArgumentOutOfRangeException Caught")
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["ArgumentOutOfRangeException Caught"]);
}

#[test]
fn test_vb_observable_collection_struct_elements() {
    let src = r#"
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
        Console.WriteLine(addedPt.X & "," & addedPt.Y)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["5,10"]);
}

#[test]
fn test_vb_observable_collection_unsubscribing_collection_changed() {
    let src = r#"
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
        Console.WriteLine(count)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1"]);
}

#[test]
fn test_vb_observable_collection_constructor_existing_list() {
    let src = r#"
Imports System.Collections.Generic
Imports System.Collections.ObjectModel

Module Program
    Sub Main()
        Dim initialList As New List(Of String) From {"One", "Two"}
        Dim col As New ObservableCollection(Of String)(initialList)
        Console.WriteLine(col.Count & "|" & String.Join(",", col))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2|One,Two"]);
}

#[test]
fn test_vb_observable_collection_set_same_index_value_triggers_replace() {
    let src = r#"
Imports System.Collections.ObjectModel
Imports System.Collections.Specialized

Module Program
    Sub Main()
        Dim col As New ObservableCollection(Of String) From {"Same"}
        Dim fired = False
        AddHandler col.CollectionChanged, Sub(s, e)
            If e.Action = NotifyCollectionChangedAction.Replace Then fired = True
        End Sub
        col(0) = "Same"
        Console.WriteLine(fired)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_observable_collection_linq_query_projections() {
    let src = r#"
Imports System.Collections.ObjectModel
Imports System.Linq

Module Program
    Sub Main()
        Dim col As New ObservableCollection(Of Integer) From {1, 2, 3, 4, 5}
        Dim evens = col.Where(Function(n) n Mod 2 = 0).ToList()
        Console.WriteLine(String.Join(",", evens))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2,4"]);
}
