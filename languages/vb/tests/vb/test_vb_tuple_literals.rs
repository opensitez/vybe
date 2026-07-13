use super::helpers::run_vb;

#[test]
fn tuple_literals() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        ' Value Tuple literals in VB 15
        Dim t1 = (1, "A")
        Console.WriteLine(t1.Item1)
        Console.WriteLine(t1.Item2)
        
        ' Named tuple elements
        Dim t2 = (Id:=2, Name:="B")
        Console.WriteLine(t2.Id)
        Console.WriteLine(t2.Name)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["1", "A", "2", "B"]);
}
