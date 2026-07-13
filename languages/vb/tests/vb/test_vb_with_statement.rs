use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: With Statement Advanced Tests
// ═══════════════════════════════════════════════════════════

#[test]
fn with_nested() {
    let out = run_vb(
        r#"
Class Address
    Public City As String
    Public ZipCode As Integer
End Class

Class Person
    Public Name As String
    Public Location As New Address()
End Class

Module M
    Sub Main()
        Dim p As New Person()
        With p
            .Name = "Bob"
            With .Location
                .City = "New York"
                .ZipCode = 10001
            End With
        End With
        Console.WriteLine(p.Name)
        Console.WriteLine(p.Location.City)
        Console.WriteLine(p.Location.ZipCode)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Bob", "New York", "10001"]);
}

#[test]
fn with_method_call() {
    let out = run_vb(
        r#"
Class Counter
    Public Value As Integer = 0
    Public Sub Increment()
        Value = Value + 1
    End Sub
    Public Sub Add(amount As Integer)
        Value = Value + amount
    End Sub
End Class

Module M
    Sub Main()
        Dim c As New Counter()
        With c
            .Increment()
            .Add(5)
            Console.WriteLine(.Value)
        End With
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn with_expression_evaluation() {
    // Ensuring the expression in With is evaluated only once
    let out = run_vb(
        r#"
Class Box
    Public Value As Integer
End Class

Module M
    Dim evalCount As Integer = 0
    
    Function GetBox() As Box
        evalCount = evalCount + 1
        Dim b As New Box()
        b.Value = 10
        Return b
    End Function

    Sub Main()
        With GetBox()
            .Value = .Value * 2
            Console.WriteLine(.Value)
        End With
        Console.WriteLine(evalCount)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["20", "1"]);
}
