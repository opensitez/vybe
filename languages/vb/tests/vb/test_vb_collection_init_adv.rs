use super::helpers::run_vb;

#[test]
fn collection_init_list() {
    let out = run_vb(
        r#"
Imports System.Collections.Generic

Module M
    Sub Main()
        Dim list As New List(Of String) From {"A", "B", "C"}
        Console.WriteLine(list.Count)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn collection_init_dict() {
    let out = run_vb(
        r#"
Imports System.Collections.Generic

Module M
    Sub Main()
        Dim dict As New Dictionary(Of Integer, String) From {
            {1, "One"},
            {2, "Two"}
        }
        Console.WriteLine(dict(2))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Two"]);
}

#[test]
fn collection_init_custom() {
    let out = run_vb(
        r#"
Imports System.Collections
Imports System.Collections.Generic

Class MyCol
    Implements IEnumerable(Of Integer)
    
    Private items As New List(Of Integer)
    
    Public Sub Add(val As Integer)
        items.Add(val * 2)
    End Sub
    
    Public Iterator Function GetEnumerator() As IEnumerator(Of Integer) Implements IEnumerable(Of Integer).GetEnumerator
        For Each item In items
            Yield item
        Next
    End Function

    Private Iterator Function IEnumerable_GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
        For Each item In items
            Yield item
        Next
    End Function
End Class

Module M
    Sub Main()
        Dim c As New MyCol From { 1, 2, 3 }
        Dim sum = 0
        For Each i In c
            sum += i
        Next
        Console.WriteLine(sum)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["12"]);
}
