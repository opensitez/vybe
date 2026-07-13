use super::helpers::run_vb;

#[test]
fn generics_inference_method() {
    let out = run_vb(
        r#"
Module M
    Function CreateArray(Of T)(item1 As T, item2 As T) As T()
        Return {item1, item2}
    End Function

    Sub Main()
        ' Type is inferred from arguments
        Dim arr = CreateArray(10, 20)
        Console.WriteLine(arr(0))
        Console.WriteLine(arr(1))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["10", "20"]);
}

#[test]
fn generics_inference_lambda() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim process = Function(Of T)(val As T) val
        
        Console.WriteLine(process("Hello"))
        Console.WriteLine(process(42))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Hello", "42"]);
}
