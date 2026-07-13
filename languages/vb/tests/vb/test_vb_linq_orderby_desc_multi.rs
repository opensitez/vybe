use super::helpers::run_vb;

#[test]
fn linq_orderby_desc_multi() {
    let out = run_vb(
        r#"
Imports System.Linq

Class Item
    Public Id As Integer
    Public Val As Integer
End Class

Module M
    Sub Main()
        Dim items = {New Item With {.Id = 1, .Val = 10}, New Item With {.Id = 2, .Val = 10}, New Item With {.Id = 3, .Val = 5}}
        
        Dim query = From i In items
                    Order By i.Val Descending, i.Id Ascending
                    Select i.Id
                    
        For Each id In query
            Console.WriteLine(id)
        Next
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["1", "2", "3"]);
}
