use super::helpers::run_vb;

#[test]
fn conversion_keywords() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim s As String = "123"
        Dim d As Double = 45.67
        
        ' CInt converts to Integer (rounds)
        Dim i As Integer = CInt(d)
        Console.WriteLine(i) ' Rounds 45.67 to 46
        
        ' CStr converts to String
        Dim strVal = CStr(100)
        Console.WriteLine(strVal)
        
        ' CDbl converts to Double
        Dim dVal = CDbl(s)
        Console.WriteLine(dVal + 1)
        
        ' CBool converts to Boolean
        Dim bVal = CBool("True")
        Console.WriteLine(bVal)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["46", "100", "124", "True"]);
}
