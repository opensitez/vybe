use super::helpers::run_vb;

#[test]
fn uri_is_absolute_or_relative() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim absoluteUri As New Uri("https://example.com/api")
        Dim relativeUri As New Uri("/api/status", UriKind.Relative)

        Console.WriteLine(absoluteUri.IsAbsoluteUri)
        Console.WriteLine(relativeUri.IsAbsoluteUri)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn uri_combine_base_and_relative() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim base As New Uri("https://example.com/a/")
        Dim child As New Uri(base, "b/c")

        Console.WriteLine(child.AbsoluteUri)
        Console.WriteLine(child.AbsolutePath)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["https://example.com/a/b/c", "/a/b/c"]);
}

#[test]
fn uri_scheme_host_port_are_exposed() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim uri As New Uri("https://example.com:8443/path?x=1#top")
        Console.WriteLine(uri.Scheme)
        Console.WriteLine(uri.Host)
        Console.WriteLine(uri.Port)
        Console.WriteLine(uri.Fragment)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["https", "example.com", "8443", "#top"]);
}

#[test]
fn uri_query_and_path_parts_are_preserved() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim uri As New Uri("https://example.com/search?q=vb&limit=5")
        Console.WriteLine(uri.PathAndQuery)
        Console.WriteLine(uri.Query)
        Console.WriteLine(uri.AbsolutePath)
    End Sub
End Module
"#,
    );

    assert_eq!(
        out,
        vec!["/search?q=vb&limit=5", "?q=vb&limit=5", "/search"]
    );
}

#[test]
fn uri_host_and_authority_contracts() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim uri As New Uri("https://user:pass@example.com:9443/")
        Console.WriteLine(uri.Host)
        Console.WriteLine(uri.Authority)
        Console.WriteLine(uri.UserInfo)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["example.com", "example.com:9443", "user:pass"]);
}

#[test]
fn uri_from_string_to_constructor_roundtrip() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim original As Uri = New Uri("https://example.com/blog/index.html")
        Dim parsed As Uri = New Uri(original.ToString())
        Console.WriteLine(parsed = original)
        Console.WriteLine(parsed.AbsoluteUri)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "https://example.com/blog/index.html"]);
}

#[test]
fn uri_trycreate_absolute_or_relative() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim absolute As Uri = Nothing
        Dim relative As Uri = Nothing

        Console.WriteLine(Uri.TryCreate("https://example.com", UriKind.Absolute, absolute))
        Console.WriteLine(Uri.TryCreate("bad uri !", UriKind.Absolute, relative))
        Console.WriteLine(absolute IsNot Nothing)
        Console.WriteLine(relative Is Nothing)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "False", "True", "True"]);
}

#[test]
fn uri_unescape_data_string() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim encoded As String = Uri.EscapeDataString("a b")
        Dim decoded As String = Uri.UnescapeDataString(encoded)

        Console.WriteLine(encoded)
        Console.WriteLine(decoded)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["a%20b", "a b"]);
}
