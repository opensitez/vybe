use super::helpers::run_vb;

#[test]
fn null_conditional_dictionary() {
    let out = run_vb(
        r#"
Imports System.Collections.Generic

Module M
    Sub Main()
        Dim dict As Dictionary(Of String, String) = Nothing
        
        ' Null conditional dictionary indexing
        Dim val = dict?("Key")
        Console.WriteLine(val Is Nothing)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["True"]);
}
