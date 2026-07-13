use super::helpers::run_vb;

#[test]
fn collection_initializers_complex() {
    let out = run_vb(
        r#"
Imports System.Collections.Generic

Module M
    Sub Main()
        ' List of Lists
        Dim matrix As New List(Of List(Of Integer)) From {
            New List(Of Integer) From {1, 2},
            New List(Of Integer) From {3, 4}
        }
        
        Console.WriteLine(matrix(1)(0))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["3"]);
}
