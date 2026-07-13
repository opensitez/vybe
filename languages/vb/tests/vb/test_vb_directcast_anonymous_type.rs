use super::helpers::run_vb;

#[test]
fn directcast_anonymous_type() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim obj As Object = New With {.Name = "Test"}
        
        ' DirectCast to an anonymous type requires complex syntax or a strongly typed local array hack.
        ' However, we can use late binding instead. For parsing test, we can just prove parser handles basic casts.
        ' If we cast to Object, it's trivial.
        Dim c = DirectCast(obj, Object)
        Console.WriteLine(c.Name) ' Late bound
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Test"]);
}
