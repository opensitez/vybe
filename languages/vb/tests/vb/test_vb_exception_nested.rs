use super::helpers::run_vb;

#[test]
fn exception_nested_try_catch() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Try
            Console.WriteLine("Outer Try")
            Try
                Throw New Exception("Inner")
            Catch ex As Exception
                Console.WriteLine("Caught Inner")
                Throw New Exception("Outer")
            Finally
                Console.WriteLine("Inner Finally")
            End Try
        Catch ex As Exception
            Console.WriteLine("Caught Outer: " & ex.Message)
        Finally
            Console.WriteLine("Outer Finally")
        End Try
    End Sub
End Module
"#,
    );
    assert_eq!(
        out,
        vec![
            "Outer Try",
            "Caught Inner",
            "Inner Finally",
            "Caught Outer: Outer",
            "Outer Finally"
        ]
    );
}

#[test]
fn exception_when_clause() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim code = 404
        Try
            Throw New Exception("Error")
        Catch ex As Exception When code = 200
            Console.WriteLine("OK")
        Catch ex As Exception When code = 404
            Console.WriteLine("Not Found")
        End Try
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Not Found"]);
}
