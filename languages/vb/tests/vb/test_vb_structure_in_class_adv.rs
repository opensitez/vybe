use super::helpers::run_vb;

#[test]
fn structure_in_class_adv() {
    let out = run_vb(
        r#"
Class Outer
    Public Structure Inner
        Public Val As Integer
    End Structure
End Class

Module M
    Sub Main()
        Dim i As New Outer.Inner()
        i.Val = 10
        Console.WriteLine(i.Val)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["10"]);
}
