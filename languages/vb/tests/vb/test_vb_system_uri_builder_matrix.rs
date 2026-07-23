use super::helpers::run_vb;

#[test]
fn uri_builder_populates_uri_properties() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim builder As New UriBuilder()
        builder.Scheme = "https"
        builder.Host = "example.com"
        builder.Port = 443
        builder.Path = "/api/v1"
        builder.Query = "page=1"
        builder.Fragment = "section"

        Console.WriteLine(builder.Uri.Scheme)
        Console.WriteLine(builder.Uri.Host)
        Console.WriteLine(builder.Uri.Port)
        Console.WriteLine(builder.Uri.AbsolutePath)
        Console.WriteLine(builder.Uri.Query)
        Console.WriteLine(builder.Uri.Fragment)
    End Sub
End Module
"#,
    );

    assert_eq!(
        out,
        vec![
            "https",
            "example.com",
            "443",
            "/api/v1",
            "?page=1",
            "#section"
        ]
    );
}

#[test]
fn uri_builder_roundtrips_from_existing_uri() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim original As New Uri("https://example.com/blog/index.html?x=1#top")
        Dim builder As New UriBuilder(original)

        Console.WriteLine(builder.Uri.AbsoluteUri)
        Console.WriteLine(builder.Path)
        Console.WriteLine(builder.Query)
        Console.WriteLine(builder.Fragment)
    End Sub
End Module
"#,
    );

    assert_eq!(
        out,
        vec![
            "https://example.com/blog/index.html?x=1#top",
            "/blog/index.html",
            "?x=1",
            "#top"
        ]
    );
}

#[test]
fn uri_builder_user_info_and_port() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim builder As New UriBuilder()
        builder.Scheme = "https"
        builder.Host = "example.com"
        builder.UserName = "alice"
        builder.Password = "secret"
        builder.Port = 8443

        Dim built As Uri = builder.Uri
        Console.WriteLine(built.UserInfo)
        Console.WriteLine(built.Authority)
        Console.WriteLine(built.Port)
    End Sub
End Module
"#,
    );

    assert_eq!(
        out,
        vec!["alice:secret", "alice:secret@example.com:8443", "8443"]
    );
}

#[test]
fn uri_builder_empty_host_becomes_relative_uri() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim builder As New UriBuilder()
        builder.Path = "relative/path"
        builder.Query = "x=1"

        Console.WriteLine(builder.Uri.IsAbsoluteUri)
        Console.WriteLine(builder.Uri.PathAndQuery)
        Console.WriteLine(builder.Uri.ToString())
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["False", "relative/path?x=1", "relative/path?x=1"]);
}

#[test]
fn uri_builder_can_mutate_scheme_and_recompute() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim builder As New UriBuilder("https://example.com/api")
        builder.Scheme = "http"
        builder.Port = 8080
        Dim rebuilt As String = builder.Uri.ToString()

        Console.WriteLine(rebuilt.StartsWith("http://example.com:8080"))
        Console.WriteLine(rebuilt.Contains("/api"))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "True"]);
}
