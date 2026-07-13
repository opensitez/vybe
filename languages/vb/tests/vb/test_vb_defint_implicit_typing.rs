use super::helpers::run_vb;

#[test]
fn defint_implicit_typing() {
    let out = run_vb(
        r#"
' Variables starting with I through N default to Integer
DefInt I-N
' Variables starting with S default to String
DefStr S

Module M
    Sub Main()
        ' iVar starts with I, so it is an Integer implicitly
        Dim iVar = 10
        Dim nVar = 20
        Dim sVar = "Hello"
        
        Console.WriteLine(iVar.GetType().Name)
        Console.WriteLine(sVar.GetType().Name)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Int32", "String"]);
}
