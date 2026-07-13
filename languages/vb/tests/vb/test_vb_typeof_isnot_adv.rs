use super::helpers::run_vb;

#[test]
fn typeof_isnot_adv() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim obj As Object = "String"
        
        If TypeOf obj IsNot Integer Then
            Console.WriteLine("Not Integer")
        End If
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Not Integer"]);
}
