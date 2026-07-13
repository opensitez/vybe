use super::helpers::run_vb;

#[test]
fn global_namespace_targeting() {
    let out = run_vb(
        r#"
Namespace Root
    Public Class A
        Public Sub Show()
            Console.WriteLine("Root.A")
        End Sub
    End Class
End Namespace

Namespace Nested
    Public Class A
        Public Sub Show()
            Console.WriteLine("Nested.A")
        End Sub
    End Class

    Module M
        Sub Main()
            Dim obj1 As New A()
            obj1.Show()
            
            Dim obj2 As New Global.Root.A()
            obj2.Show()
        End Sub
    End Module
End Namespace
"#,
    );
    assert_eq!(out, vec!["Nested.A", "Root.A"]);
}
