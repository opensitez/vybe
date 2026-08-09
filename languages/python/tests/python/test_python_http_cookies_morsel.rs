#![allow(non_snake_case)]
use super::helpers::run_python;

// http.cookies — SimpleCookie, Morsel, CookieError, Set-Cookie header generation, attributes (max-age, expires, path, domain, secure, httponly, samesite), raw cookie string parsing

#[test]
fn test_http_cookies_simple_cookie_set_and_get() {
    let out = run_python(
        r#"
from http.cookies import SimpleCookie
c = SimpleCookie()
c["session_id"] = "xyz12345"
print(c["session_id"].value)
"#,
    );
    assert_eq!(out, vec!["xyz12345"]);
}

#[test]
fn test_http_cookies_output_header_string() {
    let out = run_python(
        r#"
from http.cookies import SimpleCookie
c = SimpleCookie()
c["user"] = "bob"
print(c.output().strip())
"#,
    );
    assert_eq!(out, vec!["Set-Cookie: user=bob"]);
}

#[test]
fn test_http_cookies_morsel_attributes_path_and_domain() {
    let out = run_python(
        r#"
from http.cookies import SimpleCookie
c = SimpleCookie()
c["token"] = "abc"
c["token"]["path"] = "/api"
c["token"]["domain"] = "example.com"
out = c.output()
print("path=/api" in out)
print("domain=example.com" in out)
"#,
    );
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_http_cookies_morsel_secure_and_httponly() {
    let out = run_python(
        r#"
from http.cookies import SimpleCookie
c = SimpleCookie()
c["auth"] = "secret"
c["auth"]["secure"] = True
c["auth"]["httponly"] = True
out = c.output()
print("Secure" in out or "secure" in out)
print("HttpOnly" in out or "httponly" in out)
"#,
    );
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_http_cookies_morsel_max_age_and_expires() {
    let out = run_python(
        r#"
from http.cookies import SimpleCookie
c = SimpleCookie()
c["visited"] = "1"
c["visited"]["max-age"] = 3600
out = c.output()
print("Max-Age=3600" in out or "max-age=3600" in out)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_http_cookies_samesite_attribute() {
    let out = run_python(
        r#"
from http.cookies import SimpleCookie
c = SimpleCookie()
c["theme"] = "dark"
c["theme"]["samesite"] = "Lax"
out = c.output()
print("SameSite=Lax" in out or "samesite=Lax" in out)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_http_cookies_parse_cookie_header_string() {
    let out = run_python(
        r#"
from http.cookies import SimpleCookie
raw = "user=alice; session=9876; theme=light"
c = SimpleCookie()
c.load(raw)
print(c["user"].value)
print(c["session"].value)
print(c["theme"].value)
"#,
    );
    assert_eq!(out, vec!["alice", "9876", "light"]);
}

#[test]
fn test_http_cookies_js_output_format() {
    let out = run_python(
        r#"
from http.cookies import SimpleCookie
c = SimpleCookie()
c["key"] = "val"
print(c.js_output().strip())
"#,
    );
    assert_eq!(out, vec!["document.cookie = \"key=val;\";"]);
}

#[test]
fn test_http_cookies_value_with_quotes() {
    let out = run_python(
        r#"
from http.cookies import SimpleCookie
c = SimpleCookie()
c["quoted"] = '"hello world"'
print(c["quoted"].value)
"#,
    );
    assert_eq!(out, vec!["hello world"]);
}

#[test]
fn test_http_cookies_value_decoding_coded_value() {
    let out = run_python(
        r#"
from http.cookies import SimpleCookie
c = SimpleCookie()
c["test"] = "foo bar"
print(c["test"].coded_value)
"#,
    );
    assert_eq!(out, vec!["\"foo bar\""]);
}

#[test]
fn test_http_cookies_delete_cookie_setting_max_age_zero() {
    let out = run_python(
        r#"
from http.cookies import SimpleCookie
c = SimpleCookie()
c["del_me"] = ""
c["del_me"]["max-age"] = 0
out = c.output()
print("del_me=" in out)
print("Max-Age=0" in out or "max-age=0" in out)
"#,
    );
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_http_cookies_cookie_error_on_invalid_key() {
    let out = run_python(
        r#"
from http.cookies import SimpleCookie, CookieError
c = SimpleCookie()
try:
    c["invalid key with spaces"] = "val"
except CookieError:
    print("CookieError")
"#,
    );
    assert_eq!(out, vec!["CookieError"]);
}

#[test]
fn test_http_cookies_multiple_cookies_output() {
    let out = run_python(
        r#"
from http.cookies import SimpleCookie
c = SimpleCookie()
c["c1"] = "v1"
c["c2"] = "v2"
lines = sorted(c.output().strip().split("\n"))
print(lines[0])
print(lines[1])
"#,
    );
    assert_eq!(out, vec!["Set-Cookie: c1=v1", "Set-Cookie: c2=v2"]);
}

#[test]
fn test_http_cookies_morsel_keys() {
    let out = run_python(
        r#"
from http.cookies import SimpleCookie
c = SimpleCookie()
c["k"] = "v"
m = c["k"]
print("path" in m)
print("domain" in m)
print("max-age" in m)
print("secure" in m)
print("httponly" in m)
"#,
    );
    assert_eq!(out, vec!["True", "True", "True", "True", "True"]);
}

#[test]
fn test_http_cookies_parse_multiple_semicolon_separated() {
    let out = run_python(
        r#"
from http.cookies import SimpleCookie
c = SimpleCookie()
c.load("a=1; b=2; c=3")
print(len(c))
print(list(c.keys()))
"#,
    );
    assert_eq!(out, vec!["3", "['a', 'b', 'c']"]);
}

#[test]
fn test_http_cookies_morsel_set_value_directly() {
    let out = run_python(
        r#"
from http.cookies import SimpleCookie
c = SimpleCookie()
c["lang"] = "en"
c["lang"].set("lang", "fr", "fr")
print(c["lang"].value)
"#,
    );
    assert_eq!(out, vec!["fr"]);
}

#[test]
fn test_http_cookies_clear_cookie() {
    let out = run_python(
        r#"
from http.cookies import SimpleCookie
c = SimpleCookie()
c["x"] = "1"
c.clear()
print(len(c))
"#,
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn test_http_cookies_load_dict_mapping() {
    let out = run_python(
        r#"
from http.cookies import SimpleCookie
c = SimpleCookie({"k1": "v1", "k2": "v2"})
print(c["k1"].value)
print(c["k2"].value)
"#,
    );
    assert_eq!(out, vec!["v1", "v2"]);
}

#[test]
fn test_http_cookies_output_sep_header() {
    let out = run_python(
        r#"
from http.cookies import SimpleCookie
c = SimpleCookie()
c["a"] = "1"
c["b"] = "2"
print(c.output(sep="; "))
"#,
    );
    assert_eq!(out, vec!["Set-Cookie: a=1; Set-Cookie: b=2"]);
}

#[test]
fn test_http_cookies_morsel_output_string() {
    let out = run_python(
        r#"
from http.cookies import SimpleCookie
c = SimpleCookie()
c["id"] = "100"
c["id"]["httponly"] = True
print("HttpOnly" in c["id"].OutputString() or "httponly" in c["id"].OutputString())
"#,
    );
    assert_eq!(out, vec!["True"]);
}
