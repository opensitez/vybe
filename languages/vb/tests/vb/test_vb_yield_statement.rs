use super::helpers::run_vb;

#[test]
fn yield_statement_iterator() {
    let out = run_vb(
        r#"
Imports System.Collections.Generic

Module M
    Iterator Function GetNumbers() As IEnumerable(Of Integer)
        Yield 10
        Yield 20
        Yield 30
    End Function

    Sub Main()
        For Each num In GetNumbers()
            Console.WriteLine(num)
        Next
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["10", "20", "30"]);
}

#[test]
fn exit_function_iterator() {
    let out = run_vb(
        r#"
Imports System.Collections.Generic

Module M
    Iterator Function GetNumbersUntil() As IEnumerable(Of Integer)
        Yield 1
        Exit Function ' Terminates iterator
        Yield 2
    End Function

    Sub Main()
        For Each num In GetNumbersUntil()
            Console.WriteLine(num)
        Next
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["1"]);
}
