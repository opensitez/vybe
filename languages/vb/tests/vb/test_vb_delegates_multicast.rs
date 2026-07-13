use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Delegates (Multicast and Return Values)
// ═══════════════════════════════════════════════════════════

#[test]
fn delegates_multicast_combine() {
    let out = run_vb(
        r#"
Delegate Sub LogAction(msg As String)

Module M
    Sub LogToConsole(msg As String)
        Console.WriteLine("Console: " & msg)
    End Sub
    
    Sub LogToFile(msg As String)
        Console.WriteLine("File: " & msg)
    End Sub

    Sub Main()
        Dim d1 As LogAction = AddressOf LogToConsole
        Dim d2 As LogAction = AddressOf LogToFile
        
        ' Multicast delegate combination
        Dim d3 As LogAction = CType([Delegate].Combine(d1, d2), LogAction)
        
        d3("Test")
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Console: Test", "File: Test"]);
}

#[test]
fn delegates_return_value_multicast() {
    let out = run_vb(
        r#"
Delegate Function Calc(x As Integer) As Integer

Module M
    Function DoubleIt(x As Integer) As Integer
        Return x * 2
    End Function
    
    Function TripleIt(x As Integer) As Integer
        Return x * 3
    End Function

    Sub Main()
        Dim d1 As Calc = AddressOf DoubleIt
        Dim d2 As Calc = AddressOf TripleIt
        
        Dim d3 As Calc = CType([Delegate].Combine(d1, d2), Calc)
        
        ' When a multicast delegate returns a value, it returns the value from the last method invoked.
        Dim result As Integer = d3(5)
        Console.WriteLine(result)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["15"]);
}
