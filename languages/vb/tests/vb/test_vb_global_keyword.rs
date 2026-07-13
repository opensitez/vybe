use super::helpers::run_vb;

#[test]
fn global_keyword_namespace() {
    let out = run_vb(
        r#"
Namespace System
    Class Console
        Public Shared Sub WriteLine(s As String)
            ' Shadowing the real System.Console
        End Sub
    End Class
End Namespace

Module M
    Sub Main()
        ' Using Global allows escaping the local namespace shadowing to hit the root
        Global.System.Console.WriteLine("Hit Root")
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Hit Root"]);
}
