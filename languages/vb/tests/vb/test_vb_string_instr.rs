use super::helpers::run_vb;

#[test]
fn string_instr() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim s As String = "abracadabra"
        
        ' InStr (start_pos, string1, string2)
        ' 1-based index returns
        Console.WriteLine(InStr(1, s, "a"))
        Console.WriteLine(InStr(2, s, "a"))
        
        ' InStrRev searches from right to left
        Console.WriteLine(InStrRev(s, "a"))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["1", "4", "11"]);
}
