use super::helpers::run_vb;

#[test]
fn my_namespace_parsing() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        ' My is a virtual namespace in VB.NET
        ' Usually includes My.Application, My.Computer, My.User
        
        ' Just checking compiler support for 'My' namespace 
        ' (availability of properties depends on the framework version)
        Dim b As Boolean = True
        If Not b Then
            Console.WriteLine(My.Computer.Name)
            Console.WriteLine(My.Application.Info.Title)
            Console.WriteLine(My.User.Name)
        End If
        Console.WriteLine("My Namespace Parsed")
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["My Namespace Parsed"]);
}
