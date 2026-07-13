use super::helpers::run_vb;

#[test]
fn iterator_try_finally() {
    let out = run_vb(
        r#"
Imports System.Collections.Generic

Module M
    Iterator Function Generate() As IEnumerable(Of Integer)
        Try
            Yield 1
            Yield 2
        Finally
            Console.WriteLine("Cleanup")
        End Try
    End Function

    Sub Main()
        For Each num In Generate()
            Console.WriteLine(num)
        Next
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["1", "2", "Cleanup"]);
}
