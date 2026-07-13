use super::helpers::run_vb;

#[test]
fn notinheritable_classes() {
    let out = run_vb(
        r#"
' NotInheritable prevents other classes from inheriting from this class (like sealed in C#)
NotInheritable Class FinalClass
    Public Sub Print()
        Console.WriteLine("Final")
    End Sub
End Class

Module M
    Sub Main()
        Dim f As New FinalClass()
        f.Print()
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Final"]);
}
