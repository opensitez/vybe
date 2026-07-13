use super::helpers::run_vb;

#[test]
fn dictionary_init_nested() {
    let out = run_vb(
        r#"
Imports System.Collections.Generic

Module M
    Sub Main()
        Dim dict As New Dictionary(Of String, Object) From {
            {"A", New With {.Value = 1}},
            {"B", New With {.Value = 2}}
        }
        
        ' Late binding used to access .Value
        Console.WriteLine(dict("A").Value)
        Console.WriteLine(dict("B").Value)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["1", "2"]);
}
