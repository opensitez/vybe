use super::helpers::run_vb;

#[test]
fn info_isnumeric_isdate() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        ' IsNumeric
        Console.WriteLine(IsNumeric("123"))
        Console.WriteLine(IsNumeric("12.34"))
        Console.WriteLine(IsNumeric("abc"))
        
        ' IsDate
        Console.WriteLine(IsDate("2023-01-01"))
        Console.WriteLine(IsDate("Not A Date"))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["True", "True", "False", "True", "False"]);
}
