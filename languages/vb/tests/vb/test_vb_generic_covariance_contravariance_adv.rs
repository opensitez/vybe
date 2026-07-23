use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Generic Variance (Out / Covariance & In / Contravariance)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_generic_covariance_out_modifier() {
    let src = r#"
Public Interface IReadOnlyContainer(Of Out T)
    Function GetItem() As T
End Interface

Class StringContainer
    Implements IReadOnlyContainer(Of String)
    Public Function GetItem() As String Implements IReadOnlyContainer(Of String).GetItem
        Return "CovariantString"
    End Function
End Class

Module Program
    Sub Main()
        Dim strCont As IReadOnlyContainer(Of String) = New StringContainer()
        Dim objCont As IReadOnlyContainer(Of Object) = strCont ' Covariance assignment
        Console.WriteLine(objCont.GetItem())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["CovariantString"]);
}

#[test]
fn test_vb_generic_contravariance_in_modifier() {
    let src = r#"
Public Interface IItemConsumer(Of In T)
    Sub Consume(item As T)
End Interface

Class ObjectConsumer
    Implements IItemConsumer(Of Object)
    Public Sub Consume(item As Object) Implements IItemConsumer(Of Object).Consume
        Console.WriteLine("Consumed: " & item.ToString())
    End Sub
End Class

Module Program
    Sub Main()
        Dim objCons As IItemConsumer(Of Object) = New ObjectConsumer()
        Dim strCons As IItemConsumer(Of String) = objCons ' Contravariance assignment
        strCons.Consume("ContravariantValue")
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Consumed: ContravariantValue"]);
}
