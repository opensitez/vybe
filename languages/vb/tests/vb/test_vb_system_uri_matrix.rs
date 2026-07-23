use super::helpers::run_vb;

#[test]
fn uri_parses_components() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim u As New Uri("https://example.com:8443/app/index.html?x=1#top")
        Console.WriteLine(u.Scheme)
        Console.WriteLine(u.Host)
        Console.WriteLine(u.Port)
        Console.WriteLine(u.AbsolutePath)
        Console.WriteLine(u.Query)
        Console.WriteLine(u.Fragment)
    End Sub
End Module
"#,
    );

    assert_eq!(
        out,
        vec![
            "https",
            "example.com",
            "8443",
            "/app/index.html",
            "?x=1",
            "#top"
        ]
    );
}

#[test]
fn uri_combines_base_and_relative() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim baseUri As New Uri("https://example.com/root/")
        Dim child As New Uri(baseUri, "sub/path")
        Console.WriteLine(child.AbsolutePath)
        Console.WriteLine(child.Host)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["/root/sub/path", "example.com"]);
}

#[test]
fn uri_escape_and_unescape_roundtrip() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim encoded As String = Uri.EscapeDataString("a b/c")
        Dim decoded As String = Uri.UnescapeDataString(encoded)
        Console.WriteLine(encoded.Contains("%20"))
        Console.WriteLine(decoded)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "a b/c"]);
}

#[test]
fn uri_is_well_formed_checks_valid_inputs() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Console.WriteLine(Uri.IsWellFormedUriString("https://example.com/path", UriKind.Absolute))
        Console.WriteLine(Uri.IsWellFormedUriString("@@@", UriKind.Absolute))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn uri_try_create_with_byref_output() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim parsed As Uri = Nothing
        Dim ok As Boolean = Uri.TryCreate("https://example.com", UriKind.Absolute, parsed)
        Console.WriteLine(ok)
        Console.WriteLine(parsed IsNot Nothing)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn uri_base_of_works_for_child_resources() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim root As New Uri("https://example.com/api/")
        Dim child As New Uri("https://example.com/api/users/1")
        Console.WriteLine(root.IsBaseOf(child))
        Console.WriteLine(root.IsBaseOf(New Uri("https://example.net")))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn uri_relative_path_from_base() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim root As New Uri("https://example.com/root/a/b/")
        Dim rel As Uri = root.MakeRelativeUri(New Uri("https://example.com/root/a/b/c/d.txt"))
        Console.WriteLine(rel.ToString())
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["c/d.txt"]);
}

#[test]
fn uri_local_path_for_file_uri() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim u As New Uri("file:///tmp/sample.txt")
        Console.WriteLine(u.Scheme)
        Console.WriteLine(u.IsFile)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["file", "True"]);
}

#[test]
fn uri_host_name_type_reports_dns_or_ip() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim u1 As New Uri("https://example.com")
        Dim u2 As New Uri("https://127.0.0.1")
        Console.WriteLine(u1.HostNameType.ToString())
        Console.WriteLine(u2.HostNameType.ToString())
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["Dns", "IPv4"]);
}
