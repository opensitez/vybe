use super::helpers::run_vb;

#[test]
fn class_in_structure_adv() {
    let out = run_vb(
        r#"
Structure Outer
    Public Class Inner
        Public Sub Run()
            Console.WriteLine("InnerClass")
        End Sub
    End Class
End Structure

Module M
    Sub Main()
        Dim i As New Outer.Inner()
        i.Run()
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["InnerClass"]);
}
