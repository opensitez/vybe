use super::helpers::run_vb;

#[test]
fn interface_auto_properties() {
    let out = run_vb(
        r#"
Interface IData
    ' Properties in interfaces don't need Get/Set explicitly if they are auto-implemented style
    Property Value As Integer
End Interface

Class Data
    Implements IData
    
    Public Property Value As Integer Implements IData.Value
End Class

Module M
    Sub Main()
        Dim d As IData = New Data()
        d.Value = 100
        Console.WriteLine(d.Value)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["100"]);
}
