use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Global Namespace Access
// ═══════════════════════════════════════════════════════════

#[test]
fn global_namespace_access() {
    let out = run_vb(
        r#"
Namespace MyProject.Utils
    Class Logger
        Public Sub Log(msg As String)
            Console.WriteLine("Utils Logger: " & msg)
        End Sub
    End Class
End Namespace

Class Logger
    Public Sub Log(msg As String)
        Console.WriteLine("Global Logger: " & msg)
    End Sub
End Class

Module M
    Sub Main()
        ' Accessing global namespace using the Global keyword
        Dim gLog As New Global.Logger()
        gLog.Log("Hello")
        
        Dim uLog As New Global.MyProject.Utils.Logger()
        uLog.Log("Hello")
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Global Logger: Hello", "Utils Logger: Hello"]);
}
