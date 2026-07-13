use super::helpers::run_vb;

#[test]
fn system_guid_advanced() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim g1 As Guid = Guid.NewGuid()
        Dim g2 As Guid = Guid.NewGuid()
        
        Console.WriteLine(g1.Equals(g2))
        
        Dim strGuid As String = "d87a74a4-5694-4d8b-a3ed-3085794711f1"
        Dim parsedGuid As Guid
        If Guid.TryParse(strGuid, parsedGuid) Then
            Console.WriteLine("Parsed")
        End If
        
        Console.WriteLine(parsedGuid.ToString("D").ToLower())
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["False", "Parsed", "d87a74a4-5694-4d8b-a3ed-3085794711f1"]);
}
