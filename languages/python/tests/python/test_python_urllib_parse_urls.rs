use super::helpers::run_python;

// urllib.parse — urlparse, urlunparse, urljoin, urlsplit, urlunsplit, parse_qs, parse_qsl, quote, unquote, quote_plus, unquote_plus

#[test]
fn test_urllib_parse_urlparse_components() {
    let out = run_python(
        r#"
from urllib.parse import urlparse
u = urlparse("https://user:pass@example.com:8080/path/to/resource?query=val#frag")
print(u.scheme)
print(u.netloc)
print(u.path)
print(u.query)
print(u.fragment)
print(u.port)
"#,
    );
    assert_eq!(
        out,
        vec![
            "https",
            "user:pass@example.com:8080",
            "/path/to/resource",
            "query=val",
            "frag",
            "8080"
        ]
    );
}

#[test]
fn test_urllib_parse_urlunparse_roundtrip() {
    let out = run_python(
        r#"
from urllib.parse import urlparse, urlunparse
url = "https://example.com/search?q=rust#top"
parsed = urlparse(url)
reconstructed = urlunparse(parsed)
print(reconstructed == url)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_urllib_parse_urljoin_relative_paths() {
    let out = run_python(
        r#"
from urllib.parse import urljoin
base = "https://example.com/docs/index.html"
print(urljoin(base, "guide.html"))
print(urljoin(base, "../about.html"))
print(urljoin(base, "/api/v1"))
"#,
    );
    assert_eq!(
        out,
        vec![
            "https://example.com/docs/guide.html",
            "https://example.com/about.html",
            "https://example.com/api/v1"
        ]
    );
}

#[test]
fn test_urllib_parse_urlsplit_without_params() {
    let out = run_python(
        r#"
from urllib.parse import urlsplit
s = urlsplit("http://example.com/page;param?a=1#section")
print(s.path)
print(s.query)
print(s.fragment)
"#,
    );
    assert_eq!(out, vec!["/page;param", "a=1", "section"]);
}

#[test]
fn test_urllib_parse_urlunsplit() {
    let out = run_python(
        r#"
from urllib.parse import urlsplit, urlunsplit
s = urlsplit("http://example.com/path?key=val#tag")
print(urlunsplit(s))
"#,
    );
    assert_eq!(out, vec!["http://example.com/path?key=val#tag"]);
}

#[test]
fn test_urllib_parse_parse_qs_multivalue() {
    let out = run_python(
        r#"
from urllib.parse import parse_qs
qs = "tag=python&tag=rust&category=coding"
d = parse_qs(qs)
print(d["tag"])
print(d["category"])
"#,
    );
    assert_eq!(out, vec!["['python', 'rust']", "['coding']"]);
}

#[test]
fn test_urllib_parse_parse_qsl_tuples() {
    let out = run_python(
        r#"
from urllib.parse import parse_qsl
qs = "a=1&b=2&a=3"
pairs = parse_qsl(qs)
print(pairs)
"#,
    );
    assert_eq!(out, vec!["[('a', '1'), ('b', '2'), ('a', '3')]"]);
}

#[test]
fn test_urllib_parse_quote_escapes_special_chars() {
    let out = run_python(
        r#"
from urllib.parse import quote
print(quote("hello world!"))
print(quote("foo/bar"))
"#,
    );
    assert_eq!(out, vec!["hello%20world%21", "foo/bar"]);
}

#[test]
fn test_urllib_parse_quote_safe_arg() {
    let out = run_python(
        r#"
from urllib.parse import quote
print(quote("foo/bar", safe=""))
"#,
    );
    assert_eq!(out, vec!["foo%2Fbar"]);
}

#[test]
fn test_urllib_parse_unquote_decodes_percent() {
    let out = run_python(
        r#"
from urllib.parse import unquote
print(unquote("hello%20world%21"))
"#,
    );
    assert_eq!(out, vec!["hello world!"]);
}

#[test]
fn test_urllib_parse_quote_plus_spaces_as_plus() {
    let out = run_python(
        r#"
from urllib.parse import quote_plus
print(quote_plus("hello world"))
"#,
    );
    assert_eq!(out, vec!["hello+world"]);
}

#[test]
fn test_urllib_parse_unquote_plus() {
    let out = run_python(
        r#"
from urllib.parse import unquote_plus
print(unquote_plus("hello+world%21"))
"#,
    );
    assert_eq!(out, vec!["hello world!"]);
}

#[test]
fn test_urllib_parse_urlencode_dict() {
    let out = run_python(
        r#"
from urllib.parse import urlencode
params = {"name": "Alice", "city": "New York"}
print(urlencode(params))
"#,
    );
    assert_eq!(out, vec!["name=Alice&city=New+York"]);
}

#[test]
fn test_urllib_parse_urlencode_doseq() {
    let out = run_python(
        r#"
from urllib.parse import urlencode
params = {"id": [1, 2, 3]}
print(urlencode(params, doseq=True))
"#,
    );
    assert_eq!(out, vec!["id=1&id=2&id=3"]);
}

#[test]
fn test_urllib_parse_urldefrag_removes_fragment() {
    let out = run_python(
        r#"
from urllib.parse import urldefrag
defragged, frag = urldefrag("https://example.com/page.html#heading")
print(defragged)
print(frag)
"#,
    );
    assert_eq!(out, vec!["https://example.com/page.html", "heading"]);
}

#[test]
fn test_urllib_parse_parse_qs_keep_blank_values() {
    let out = run_python(
        r#"
from urllib.parse import parse_qs
qs = "a=&b=1"
print(parse_qs(qs, keep_blank_values=True))
print(parse_qs(qs, keep_blank_values=False))
"#,
    );
    assert_eq!(out, vec!["{'a': [''], 'b': ['1']}", "{'b': ['1']}"]);
}

#[test]
fn test_urllib_parse_username_and_password() {
    let out = run_python(
        r#"
from urllib.parse import urlparse
u = urlparse("ftp://admin:secret123@files.org:21")
print(u.username)
print(u.password)
print(u.hostname)
"#,
    );
    assert_eq!(out, vec!["admin", "secret123", "files.org"]);
}

#[test]
fn test_urllib_parse_quote_bytes_input() {
    let out = run_python(
        r#"
from urllib.parse import quote_from_bytes
print(quote_from_bytes(b"hello world\xff"))
"#,
    );
    assert_eq!(out, vec!["hello%20world%FF"]);
}

#[test]
fn test_urllib_parse_unquote_to_bytes() {
    let out = run_python(
        r#"
from urllib.parse import unquote_to_bytes
b = unquote_to_bytes("hello%20world%FF")
print(isinstance(b, bytes))
print(b.endswith(b"\xff"))
"#,
    );
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_urllib_parse_urlparse_ipv6_host() {
    let out = run_python(
        r#"
from urllib.parse import urlparse
u = urlparse("http://[::1]:8080/index")
print(u.hostname)
print(u.port)
"#,
    );
    assert_eq!(out, vec!["::1", "8080"]);
}
