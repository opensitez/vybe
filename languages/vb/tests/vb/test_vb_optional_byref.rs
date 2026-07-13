use super::helpers::run_vb;

#[test]
fn optional_byref() {
    let out = run_vb(
        r#"
Module M
    ' Optional ByRef parameter
    Sub Process(Optional ByRef val As Integer = 5)
        val += 10
        Console.WriteLine(val)
    End Sub

    Sub Main()
        Process()
        
        Dim x = 100
        Process(x)
        Console.WriteLine(x)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["15", "110", "110"]);
}
