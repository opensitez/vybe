use super::helpers::run_vb;

#[test]
fn system_collections_concurrent_dict() {
    let out = run_vb(
        r#"
Imports System.Collections.Concurrent

Module M
    Sub Main()
        Dim cd As New ConcurrentDictionary(Of Integer, String)()
        
        cd.TryAdd(1, "One")
        cd.TryAdd(2, "Two")
        
        Dim val As String = Nothing
        If cd.TryGetValue(1, val) Then
            Console.WriteLine(val)
        End If
        
        ' Update
        cd.AddOrUpdate(1, "NewOne", Function(k, oldVal) "NewOne")
        Console.WriteLine(cd(1))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["One", "NewOne"]);
}

#[test]
fn system_collections_concurrent_queue() {
    let out = run_vb(
        r#"
Imports System.Collections.Concurrent

Module M
    Sub Main()
        Dim cq As New ConcurrentQueue(Of Integer)()
        
        cq.Enqueue(10)
        cq.Enqueue(20)
        
        Dim result As Integer
        If cq.TryDequeue(result) Then
            Console.WriteLine(result)
        End If
        
        Console.WriteLine(cq.Count)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["10", "1"]);
}
