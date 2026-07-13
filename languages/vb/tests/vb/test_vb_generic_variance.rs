use super::helpers::run_vb;

#[test]
fn generic_variance_in_out() {
    let out = run_vb(
        r#"
' Out defines covariance, In defines contravariance
Interface IProducer(Of Out T)
    Function Produce() As T
End Interface

Interface IConsumer(Of In T)
    Sub Consume(item As T)
End Interface

Class StringProducer
    Implements IProducer(Of String)
    
    Public Function Produce() As String Implements IProducer(Of String).Produce
        Return "Hello"
    End Function
End Class

Module M
    Sub Main()
        ' Covariance: Assigning a more specific generic type to a less specific one
        Dim p As IProducer(Of Object) = New StringProducer()
        Console.WriteLine(p.Produce())
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Hello"]);
}
