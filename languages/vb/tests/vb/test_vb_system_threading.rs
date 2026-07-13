use super::helpers::run_vb;

#[test]
fn system_threading_thread() {
    let out = run_vb(
        r#"
Imports System.Threading

Module M
    Sub Main()
        Dim t As Thread = Thread.CurrentThread
        Console.WriteLine(t IsNot Nothing)
        Console.WriteLine(t.IsAlive)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn system_threading_interlocked() {
    let out = run_vb(
        r#"
Imports System.Threading

Module M
    Sub Main()
        Dim count As Integer = 0
        Interlocked.Increment(count)
        Console.WriteLine(count)
        
        Interlocked.Add(count, 5)
        Console.WriteLine(count)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["1", "6"]);
}
