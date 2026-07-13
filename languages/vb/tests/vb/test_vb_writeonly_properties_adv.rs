use super::helpers::run_vb;

#[test]
fn writeonly_properties_adv() {
    let out = run_vb(
        r#"
Class Logger
    Private _lastMsg As String
    
    ' WriteOnly property
    Public WriteOnly Property Message As String
        Set(value As String)
            _lastMsg = value
            Console.WriteLine("Logged: " & value)
        End Set
    End Property
    
    Public Function GetLast() As String
        Return _lastMsg
    End Function
End Class

Module M
    Sub Main()
        Dim l As New Logger()
        l.Message = "Start"
        l.Message = "End"
        
        Console.WriteLine(l.GetLast())
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Logged: Start", "Logged: End", "End"]);
}
