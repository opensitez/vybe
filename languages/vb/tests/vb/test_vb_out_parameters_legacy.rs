use super::helpers::run_vb;

#[test]
fn out_parameters_legacy() {
    let out = run_vb(
        r#"
Imports System.Runtime.InteropServices

Module M
    ' Legacy way of defining Out parameters in VB
    Sub GetValues(<Out> ByRef a As Integer, <Out> ByRef b As String)
        a = 100
        b = "Data"
    End Sub

    Sub Main()
        Dim a As Integer
        Dim b As String = Nothing
        GetValues(a, b)
        Console.WriteLine(a)
        Console.WriteLine(b)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["100", "Data"]);
}
