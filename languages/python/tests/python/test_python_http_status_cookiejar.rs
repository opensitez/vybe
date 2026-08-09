use super::helpers::run_python;

// http.HTTPStatus, http.cookiejar, http.client response parsing

#[test]
fn test_http_status_ok_value_and_phrase() {
    let out = run_python(
        r#"
from http import HTTPStatus
s = HTTPStatus.OK
print(s.value)
print(s.phrase)
"#,
    );
    assert_eq!(out, vec!["200", "OK"]);
}

#[test]
fn test_http_status_not_found() {
    let out = run_python(
        r#"
from http import HTTPStatus
s = HTTPStatus.NOT_FOUND
print(s.value)
print(s.phrase)
"#,
    );
    assert_eq!(out, vec!["404", "Not Found"]);
}

#[test]
fn test_http_status_internal_server_error() {
    let out = run_python(
        r#"
from http import HTTPStatus
s = HTTPStatus.INTERNAL_SERVER_ERROR
print(s.value)
print(s.phrase)
"#,
    );
    assert_eq!(out, vec!["500", "Internal Server Error"]);
}

#[test]
fn test_http_status_created() {
    let out = run_python(
        r#"
from http import HTTPStatus
s = HTTPStatus.CREATED
print(s.value)
"#,
    );
    assert_eq!(out, vec!["201"]);
}

#[test]
fn test_http_status_is_1xx_informational() {
    let out = run_python(
        r#"
from http import HTTPStatus
print(HTTPStatus.CONTINUE.is_informational)
print(HTTPStatus.OK.is_informational)
"#,
    );
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn test_http_status_is_2xx_success() {
    let out = run_python(
        r#"
from http import HTTPStatus
print(HTTPStatus.OK.is_success)
print(HTTPStatus.NOT_FOUND.is_success)
"#,
    );
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn test_http_status_is_3xx_redirect() {
    let out = run_python(
        r#"
from http import HTTPStatus
print(HTTPStatus.MOVED_PERMANENTLY.is_redirection)
print(HTTPStatus.OK.is_redirection)
"#,
    );
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn test_http_status_is_4xx_client_error() {
    let out = run_python(
        r#"
from http import HTTPStatus
print(HTTPStatus.BAD_REQUEST.is_client_error)
print(HTTPStatus.OK.is_client_error)
"#,
    );
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn test_http_status_is_5xx_server_error() {
    let out = run_python(
        r#"
from http import HTTPStatus
print(HTTPStatus.INTERNAL_SERVER_ERROR.is_server_error)
print(HTTPStatus.OK.is_server_error)
"#,
    );
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn test_http_status_from_value() {
    let out = run_python(
        r#"
from http import HTTPStatus
s = HTTPStatus(301)
print(s.phrase)
"#,
    );
    assert_eq!(out, vec!["Moved Permanently"]);
}

#[test]
fn test_http_status_description_not_empty() {
    let out = run_python(
        r#"
from http import HTTPStatus
print(len(HTTPStatus.OK.description) > 0)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_http_cookiejar_empty_by_default() {
    let out = run_python(
        r#"
import http.cookiejar
jar = http.cookiejar.CookieJar()
print(len(list(jar)))
"#,
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn test_http_cookiejar_lwpcookiejar_save_load() {
    let out = run_python(
        r#"
import http.cookiejar, tempfile, os
f = tempfile.NamedTemporaryFile(delete=False, suffix=".txt")
f.close()
jar = http.cookiejar.LWPCookieJar(f.name)
jar.save(ignore_discard=True, ignore_expires=True)
jar2 = http.cookiejar.LWPCookieJar(f.name)
jar2.load(ignore_discard=True, ignore_expires=True)
print(len(list(jar2)) == 0)
os.unlink(f.name)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_http_status_no_content() {
    let out = run_python(
        r#"
from http import HTTPStatus
s = HTTPStatus.NO_CONTENT
print(s.value)
"#,
    );
    assert_eq!(out, vec!["204"]);
}

#[test]
fn test_http_status_unauthorized() {
    let out = run_python(
        r#"
from http import HTTPStatus
s = HTTPStatus.UNAUTHORIZED
print(s.value)
print(s.phrase)
"#,
    );
    assert_eq!(out, vec!["401", "Unauthorized"]);
}

#[test]
fn test_http_status_forbidden() {
    let out = run_python(
        r#"
from http import HTTPStatus
s = HTTPStatus.FORBIDDEN
print(s.value)
"#,
    );
    assert_eq!(out, vec!["403"]);
}

#[test]
fn test_http_status_method_not_allowed() {
    let out = run_python(
        r#"
from http import HTTPStatus
s = HTTPStatus.METHOD_NOT_ALLOWED
print(s.value)
print(s.phrase)
"#,
    );
    assert_eq!(out, vec!["405", "Method Not Allowed"]);
}

#[test]
fn test_http_status_too_many_requests() {
    let out = run_python(
        r#"
from http import HTTPStatus
s = HTTPStatus.TOO_MANY_REQUESTS
print(s.value)
"#,
    );
    assert_eq!(out, vec!["429"]);
}

#[test]
fn test_http_status_invalid_raises_value_error() {
    let out = run_python(
        r#"
from http import HTTPStatus
try:
    HTTPStatus(999)
except ValueError:
    print("ValueError")
"#,
    );
    assert_eq!(out, vec!["ValueError"]);
}

#[test]
fn test_http_status_str_representation() {
    let out = run_python(
        r#"
from http import HTTPStatus
print(str(HTTPStatus.OK))
"#,
    );
    assert_eq!(out, vec!["HTTPStatus.OK"]);
}
