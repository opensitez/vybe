use crate::helpers::run_prints;

#[test]
fn test_uri_parses_components() {
    let out = run_prints(
        r#"
        import java.net.URI

        fun main() {
            val uri = URI("https://user:pass@example.com:9443/search?q=kotlin&x=1#top")
            println(uri.scheme)
            println(uri.host)
            println(uri.port)
            println(uri.userInfo)
            println(uri.path)
            println(uri.query)
            println(uri.fragment)
            println(uri.isAbsolute)
        }
    "#,
    );
    assert_eq!(
        out,
        &[
            "https",
            "example.com",
            "9443",
            "user:pass",
            "/search",
            "q=kotlin&x=1",
            "top",
            "true"
        ]
    );
}

#[test]
fn test_uri_resolve_and_normalize() {
    let out = run_prints(
        r#"
        import java.net.URI

        fun main() {
            val base = URI("https://example.com/dir/a/b/")
            val child = base.resolve("../c/./index")
            val normalized = child.normalize()
            println(child.toString())
            println(normalized.toString())
        }
    "#,
    );
    assert_eq!(
        out,
        &[
            "https://example.com/dir/c/index",
            "https://example.com/dir/c/index"
        ]
    );
}

#[test]
fn test_uri_to_url_roundtrip() {
    let out = run_prints(
        r#"
        import java.net.URI

        fun main() {
            val uri = URI("https://example.org/resource")
            val url = uri.toURL()
            println(url.protocol)
            println(url.host)
            println(url.path)
            println(url.toURI() == uri)
        }
    "#,
    );
    assert_eq!(out, &["https", "example.org", "/resource", "true"]);
}

#[test]
fn test_url_encode_decode() {
    let out = run_prints(
        r#"
        import java.net.URLEncoder
        import java.net.URLDecoder

        fun main() {
            val encoded = URLEncoder.encode("a b/c", "UTF-8")
            val decoded = URLDecoder.decode(encoded, "UTF-8")
            println(encoded)
            println(decoded)
        }
    "#,
    );
    assert_eq!(out, &["a+b%2Fc", "a b/c"]);
}

#[test]
fn test_url_connection_protocol_metadata_only() {
    let out = run_prints(
        r#"
        import java.net.URL

        fun main() {
            val url = URL("http://localhost:8080/path?x=1")
            println(url.protocol)
            println(url.host)
            println(url.port)
            println(url.query)
            println(url.authority)
            println(url.file)
        }
    "#,
    );
    assert_eq!(
        out,
        &[
            "http",
            "localhost",
            "8080",
            "x=1",
            "localhost:8080",
            "/path?x=1"
        ]
    );
}

#[test]
fn test_uri_relativize_behavior() {
    let out = run_prints(
        r#"
        import java.net.URI

        fun main() {
            val base = URI("https://example.org/root/index")
            val target = URI("https://example.org/root/docs/page")
            val rel = base.relativize(target)
            println(rel.toString())
            println(base.resolve(rel) == target)
        }
    "#,
    );
    assert_eq!(out, &["docs/page", "true"]);
}

#[test]
fn test_uri_file_scheme_roundtrip() {
    let out = run_prints(
        r#"
        import java.net.URI

        fun main() {
            val root = java.lang.System.getProperty("java.io.tmpdir")
            val uri = URI("file", null, root, 0, "/tmp.log", null, null)
            println(uri.scheme)
            println(uri.path)
            println(uri.isAbsolute)
            println(uri.toString().startsWith("file:"))
        }
    "#,
    );
    assert_eq!(out, &["file", "/tmp.log", "true", "true"]);
}

#[test]
fn test_uri_query_parsing_manual() {
    let out = run_prints(
        r#"
        import java.net.URI

        fun main() {
            val uri = URI("https://example.com/a?left=1&right=2")
            val query = uri.query
            val parts = query.split("&").joinToString("|") { it }
            println(parts)
            println(query.contains("left=1"))
            println(query.contains("right=2"))
        }
    "#,
    );
    assert_eq!(out, &["left=1|right=2", "true", "true"]);
}
