use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: Urllib & HTTP Client Parsing — urllib.parse (urlsplit, urlunsplit, urlencode, quote, urljoin), http.client headers
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_urllib_parse_urlsplit_components() {
    let src = r#"
from urllib.parse import urlsplit

url = "https://user:pass@example.com:8080/path/to/doc?query=val#frag"
split = urlsplit(url)

print(split.scheme)
print(split.netloc)
print(split.path)
print(split.query)
print(split.fragment)
"#;
    assert_eq!(
        run_python(src),
        vec![
            "https",
            "user:pass@example.com:8080",
            "/path/to/doc",
            "query=val",
            "frag"
        ]
    );
}

#[test]
fn test_py_urllib_parse_parse_qs_parse_qsl() {
    let src = r#"
from urllib.parse import parse_qs, parse_qsl

query = "a=1&b=2&a=3"
qs = parse_qs(query)
print(qs["a"])
print(qs["b"])

qsl = parse_qsl(query)
print(qsl)
"#;
    assert_eq!(
        run_python(src),
        vec![
            "['1', '3']",
            "['2']",
            "[('a', '1'), ('b', '2'), ('a', '3')]"
        ]
    );
}

#[test]
fn test_py_urllib_parse_urlencode_dictionary() {
    let src = r#"
from urllib.parse import urlencode

params = {"name": "Alice & Bob", "city": "New York"}
encoded = urlencode(params)
print(encoded)
"#;
    assert_eq!(run_python(src), vec!["name=Alice+%26+Bob&city=New+York"]);
}

#[test]
fn test_py_urllib_parse_quote_unquote_escaping() {
    let src = r#"
from urllib.parse import quote, unquote

raw = "hello world/python&code"
q = quote(raw)
print(q)
print(unquote(q) == raw)
"#;
    assert_eq!(run_python(src), vec!["hello%20world/python%26code", "True"]);
}

#[test]
fn test_py_urllib_parse_urljoin_relative_resolution() {
    let src = r#"
from urllib.parse import urljoin

base = "https://example.com/docs/api/index.html"
print(urljoin(base, "guide.html"))
print(urljoin(base, "/v2/schema"))
print(urljoin(base, "../images/logo.png"))
"#;
    assert_eq!(
        run_python(src),
        vec![
            "https://example.com/docs/api/guide.html",
            "https://example.com/v2/schema",
            "https://example.com/docs/images/logo.png"
        ]
    );
}

#[test]
fn test_py_http_client_parse_headers_from_io() {
    let src = r#"
from http.client import parse_headers
import io

header_text = b"Content-Type: text/html\r\nContent-Length: 123\r\n\r\n"
headers = parse_headers(io.BytesIO(header_text))
print(headers["Content-Type"])
print(headers["Content-Length"])
"#;
    assert_eq!(run_python(src), vec!["text/html", "123"]);
}

#[test]
fn test_py_urllib_parse_urlunsplit_roundtrip() {
    let src = r#"
from urllib.parse import urlsplit, urlunsplit

url = "https://example.com/path?query=1"
parts = urlsplit(url)
reconstructed = urlunsplit(parts)
print(reconstructed == url)
"#;
    assert_eq!(run_python(src), vec!["True"]);
}

#[test]
fn test_py_urllib_parse_quote_plus_unquote_plus() {
    let src = r#"
from urllib.parse import quote_plus, unquote_plus

s = "space separated & query"
qp = quote_plus(s)
print(qp)
print(unquote_plus(qp) == s)
"#;
    assert_eq!(run_python(src), vec!["space+separated+%26+query", "True"]);
}

#[test]
fn test_py_http_client_responses_status_codes() {
    let src = r#"
from http.client import responses

print(responses[200])
print(responses[404])
print(responses[500])
"#;
    assert_eq!(
        run_python(src),
        vec!["OK", "Not Found", "Internal Server Error"]
    );
}

#[test]
fn test_py_urllib_request_request_object_headers() {
    let src = r#"
from urllib.request import Request

req = Request("https://example.com/api", headers={"User-Agent": "CustomApp/1.0"})
print(req.full_url)
print(req.headers["User-agent"])
"#;
    assert_eq!(
        run_python(src),
        vec!["https://example.com/api", "CustomApp/1.0"]
    );
}
