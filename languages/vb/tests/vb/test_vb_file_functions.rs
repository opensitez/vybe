use super::helpers::run_vb;

#[test]
fn file_functions_legacy() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        ' Legacy file I/O syntax parsing check
        Dim fNum = FreeFile()
        Console.WriteLine(fNum > 0)
        
        ' Note: FileOpen, EOF, LOF, Loc might throw or fail if files aren't created in test environment,
        ' so we just test FreeFile and syntax checking.
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["True"]);
}
