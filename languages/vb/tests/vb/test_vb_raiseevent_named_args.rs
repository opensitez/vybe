use super::helpers::run_vb;

#[test]
fn raiseevent_named_args() {
    let out = run_vb(
        r#"
Class Publisher
    Public Event Notify(msg As String, code As Integer)
    
    Public Sub DoNotify()
        ' RaiseEvent with named arguments
        RaiseEvent Notify(code:=100, msg:="Alert")
    End Sub
End Class

Module M
    Sub Main()
        Dim p As New Publisher()
        AddHandler p.Notify, Sub(m, c) Console.WriteLine(m & c)
        p.DoNotify()
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Alert100"]);
}
