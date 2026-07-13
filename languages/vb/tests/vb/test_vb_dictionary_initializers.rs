use super::helpers::run_vb;

#[test]
fn dictionary_initializers() {
    let out = run_vb(
        r#"
Imports System.Collections.Generic

Module M
    Sub Main()
        ' Dictionary collection initializer syntax using From { {k, v} }
        Dim dict As New Dictionary(Of Integer, String) From {
            {1, "One"},
            {2, "Two"},
            {3, "Three"}
        }
        
        Console.WriteLine(dict(2))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Two"]);
}
