use super::helpers::run_vb;

#[test]
fn yield_infinite_loop() {
    let out = run_vb(
        r#"
Imports System.Collections.Generic
Imports System.Linq

Module M
    Iterator Function Generate() As IEnumerable(Of Integer)
        Dim i = 0
        While True
            Yield i
            i += 1
        End While
    End Function

    Sub Main()
        Dim numbers = Generate().Take(3)
        For Each n In numbers
            Console.WriteLine(n)
        Next
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["0", "1", "2"]);
}
