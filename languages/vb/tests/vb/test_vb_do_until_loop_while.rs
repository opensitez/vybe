use super::helpers::run_vb;

#[test]
fn do_until_loop_while() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim i = 0
        
        ' Technically valid syntax in VB to mix conditions on Do and Loop
        Do Until i = 10
            i += 1
        Loop While i < 5
        
        Console.WriteLine(i)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["5"]);
}
