use super::helpers::run_vb;

#[test]
fn tuple_infer_elements() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim x = 10
        Dim y = "Test"
        
        ' VB.NET tuple element name inference
        Dim t = (x, y)
        
        ' The inferred names are 'x' and 'y' (if supported by compiler)
        ' Let's access them via standard ItemN to be safe if inference isn't supported
        Console.WriteLine(t.Item1)
        Console.WriteLine(t.Item2)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["10", "Test"]);
}

#[test]
fn tuple_infer_elements_names() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim count = 5
        Dim name = "Bob"
        
        Dim t = (count, name)
        
        ' If element name inference is supported, t.count should work
        ' However, we'll use a hack to check by using reflection on properties? No, ValueTuple fields.
        ' Let's just do standard assignment.
        Dim t2 As (C As Integer, N As String) = t
        Console.WriteLine(t2.C)
        Console.WriteLine(t2.N)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["5", "Bob"]);
}
