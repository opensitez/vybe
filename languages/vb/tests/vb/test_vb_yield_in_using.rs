use super::helpers::run_vb;

#[test]
fn yield_in_using() {
    let out = run_vb(
        r#"
Imports System
Imports System.Collections.Generic

Class Res
    Implements IDisposable
    Public Sub Dispose() Implements IDisposable.Dispose
        Console.WriteLine("Disposed")
    End Sub
End Class

Module M
    Iterator Function Generate() As IEnumerable(Of Integer)
        Using r As New Res()
            Yield 1
            Yield 2
        End Using
    End Function

    Sub Main()
        For Each x In Generate()
            Console.WriteLine(x)
        Next
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["1", "2", "Disposed"]);
}
