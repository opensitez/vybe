use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: ObservableCollection(Of T) & CollectionChanged Events
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_observable_collection_add_item_event() {
    let src = r#"
Imports System.Collections.ObjectModel
Imports System.Collections.Specialized

Module Program
    Sub Main()
        Dim collection As New ObservableCollection(Of String)()
        AddHandler collection.CollectionChanged, Sub(sender, e)
            Console.WriteLine("Action: " & e.Action.ToString())
            Console.WriteLine("NewItem: " & e.NewItems(0).ToString())
        End Sub

        collection.Add("FirstItem")
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Action: Add", "NewItem: FirstItem"]);
}

#[test]
fn test_vb_observable_collection_remove_item_event() {
    let src = r#"
Imports System.Collections.ObjectModel
Imports System.Collections.Specialized

Module Program
    Sub Main()
        Dim collection As New ObservableCollection(Of Integer) From {10, 20}
        AddHandler collection.CollectionChanged, Sub(sender, e)
            If e.Action = NotifyCollectionChangedAction.Remove Then
                Console.WriteLine("Removed: " & e.OldItems(0).ToString())
            End If
        End Sub

        collection.Remove(10)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Removed: 10"]);
}

#[test]
fn test_vb_observable_collection_replace_item_event() {
    let src = r#"
Imports System.Collections.ObjectModel
Imports System.Collections.Specialized

Module Program
    Sub Main()
        Dim collection As New ObservableCollection(Of String) From {"Old"}
        AddHandler collection.CollectionChanged, Sub(sender, e)
            If e.Action = NotifyCollectionChangedAction.Replace Then
                Console.WriteLine("Old: " & e.OldItems(0).ToString() & " New: " & e.NewItems(0).ToString())
            End If
        End Sub

        collection(0) = "New"
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Old: Old New: New"]);
}

#[test]
fn test_vb_observable_collection_clear_event() {
    let src = r#"
Imports System.Collections.ObjectModel
Imports System.Collections.Specialized

Module Program
    Sub Main()
        Dim collection As New ObservableCollection(Of Integer) From {1, 2, 3}
        AddHandler collection.CollectionChanged, Sub(sender, e)
            Console.WriteLine("Action: " & e.Action.ToString())
        End Sub

        collection.Clear()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Action: Reset"]);
}

#[test]
fn test_vb_observable_collection_move_item_event() {
    let src = r#"
Imports System.Collections.ObjectModel
Imports System.Collections.Specialized

Module Program
    Sub Main()
        Dim collection As New ObservableCollection(Of String) From {"A", "B", "C"}
        AddHandler collection.CollectionChanged, Sub(sender, e)
            If e.Action = NotifyCollectionChangedAction.Move Then
                Console.WriteLine("Moved from " & e.OldStartingIndex & " to " & e.NewStartingIndex)
            End If
        End Sub

        collection.Move(0, 2)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Moved from 0 to 2"]);
}

#[test]
fn test_vb_observable_collection_property_changed_event() {
    let src = r#"
Imports System.Collections.ObjectModel
Imports System.ComponentModel

Module Program
    Sub Main()
        Dim collection As New ObservableCollection(Of Integer)()
        AddHandler CType(collection, INotifyPropertyChanged).PropertyChanged, Sub(sender, e)
            Console.WriteLine("Prop: " & e.PropertyName)
        End Sub

        collection.Add(100)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Prop: Count", "Prop: Item[]"]);
}
