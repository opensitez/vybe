use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: System.UriBuilder & Uri Operations
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_uri_builder_construct_url() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim builder As New UriBuilder("https", "example.com", 8080, "api/v1")
        builder.Query = "key=value"
        Dim uri As Uri = builder.Uri
        Console.WriteLine(uri.Scheme)
        Console.WriteLine(uri.Host)
        Console.WriteLine(uri.Port)
        Console.WriteLine(uri.Query)
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["https", "example.com", "8080", "?key=value"]
    );
}
