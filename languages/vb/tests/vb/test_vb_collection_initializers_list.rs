use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Collection Initializers (List)
// ═══════════════════════════════════════════════════════════

#[test]
fn collection_initializer_list() {
    let out = run_vb(
        r#"
Imports System.Collections.Generic

Module M
    Sub Main()
        ' Collection initializer syntax
        Dim list As New List(Of String) From {"apple", "banana", "cherry"}
        
        Console.WriteLine(list.Count)
        Console.WriteLine(list(1))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["3", "banana"]);
}

#[test]
fn collection_initializer_custom_collection() {
    let out = run_vb(
        r#"
Imports System.Collections.Generic

Class MyBag
    Implements IEnumerable(Of String)
    
    Private _items As New List(Of String)()
    
    ' Required Add method for collection initializer
    Public Sub Add(item As String)
        _items.Add("My" & item)
    End Sub
    
    Public Function GetEnumerator() As IEnumerator(Of String) Implements IEnumerable(Of String).GetEnumerator
        Return _items.GetEnumerator()
    End Function
    
    Private Function IEnumerable_GetEnumerator() As System.Collections.IEnumerator Implements System.Collections.IEnumerable.GetEnumerator
        Return _items.GetEnumerator()
    End Function
End Class

Module M
    Sub Main()
        Dim bag As New MyBag From {"Cat", "Dog"}
        
        For Each b In bag
            Console.WriteLine(b)
        Next
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["MyCat", "MyDog"]);
}
