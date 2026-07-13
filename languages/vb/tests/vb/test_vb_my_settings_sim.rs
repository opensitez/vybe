use super::helpers::run_vb;

#[test]
fn my_settings_sim() {
    let out = run_vb(
        r#"
Module M
    ' Simulate My.Settings functionality
    Class SettingsClass
        Public DefaultProperty As String = "Value"
    End Class
    
    Dim Settings As New SettingsClass()
    
    Sub Main()
        Console.WriteLine(Settings.DefaultProperty)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Value"]);
}
