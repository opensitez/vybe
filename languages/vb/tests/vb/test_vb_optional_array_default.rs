use super::helpers::run_vb;

#[test]
fn optional_array_default() {
    let out = run_vb(
        r#"
Module M
    ' Arrays cannot be Optional with a default value other than Nothing.
    ' We will test parsing of Nothing as default.
    Sub PrintFirst(Optional arr() As Integer = Nothing)
        If arr IsNot Nothing Then
            Console.WriteLine(arr(0))
        Else
            Console.WriteLine("Empty")
        End If
    End Sub

    Sub Main()
        PrintFirst()
        PrintFirst({5})
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Empty", "5"]);
}
