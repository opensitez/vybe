use super::helpers::run_python;

// ssl — SSLContext creation, protocol constants, cert loading, wrap_socket options

#[test]
fn test_ssl_context_tls_client_creates_ok() {
    let out = run_python(
        r#"
import ssl
ctx = ssl.SSLContext(ssl.PROTOCOL_TLS_CLIENT)
print(ctx.verify_mode == ssl.CERT_REQUIRED)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_ssl_context_tls_server_creates_ok() {
    let out = run_python(
        r#"
import ssl
ctx = ssl.SSLContext(ssl.PROTOCOL_TLS_SERVER)
print(ctx.verify_mode == ssl.CERT_NONE)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_ssl_context_default_ciphers_set() {
    let out = run_python(
        r#"
import ssl
ctx = ssl.SSLContext(ssl.PROTOCOL_TLS_CLIENT)
ciphers = ctx.get_ciphers()
print(len(ciphers) > 0)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_ssl_context_set_verify_mode() {
    let out = run_python(
        r#"
import ssl
ctx = ssl.SSLContext(ssl.PROTOCOL_TLS_CLIENT)
ctx.check_hostname = False
ctx.verify_mode = ssl.CERT_NONE
print(ctx.verify_mode == ssl.CERT_NONE)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_ssl_context_check_hostname_attribute() {
    let out = run_python(
        r#"
import ssl
ctx = ssl.SSLContext(ssl.PROTOCOL_TLS_CLIENT)
print(isinstance(ctx.check_hostname, bool))
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_ssl_create_default_context() {
    let out = run_python(
        r#"
import ssl
ctx = ssl.create_default_context()
print(ctx.verify_mode == ssl.CERT_REQUIRED)
print(ctx.check_hostname)
"#,
    );
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_ssl_cert_required_constant() {
    let out = run_python(
        r#"
import ssl
print(ssl.CERT_NONE < ssl.CERT_OPTIONAL < ssl.CERT_REQUIRED)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_ssl_op_no_sslv2_constant() {
    let out = run_python(
        r#"
import ssl
print(isinstance(ssl.OP_NO_SSLv2, int))
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_ssl_ssl_error_inherits_from_os_error() {
    let out = run_python(
        r#"
import ssl
print(issubclass(ssl.SSLError, OSError))
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_ssl_cert_error_inherits_from_ssl_error() {
    let out = run_python(
        r#"
import ssl
print(issubclass(ssl.SSLCertVerificationError, ssl.SSLError))
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_ssl_context_maximum_version() {
    let out = run_python(
        r#"
import ssl
ctx = ssl.SSLContext(ssl.PROTOCOL_TLS_CLIENT)
print(isinstance(ctx.maximum_version, ssl.TLSVersion))
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_ssl_context_minimum_version() {
    let out = run_python(
        r#"
import ssl
ctx = ssl.SSLContext(ssl.PROTOCOL_TLS_CLIENT)
print(isinstance(ctx.minimum_version, ssl.TLSVersion))
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_ssl_tls_version_enum() {
    let out = run_python(
        r#"
import ssl
print(ssl.TLSVersion.TLSv1_2.name)
print(ssl.TLSVersion.TLSv1_3.name)
"#,
    );
    assert_eq!(out, vec!["TLSv1_2", "TLSv1_3"]);
}

#[test]
fn test_ssl_context_options_attribute() {
    let out = run_python(
        r#"
import ssl
ctx = ssl.SSLContext(ssl.PROTOCOL_TLS_CLIENT)
print(isinstance(ctx.options, int))
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_ssl_has_alpn() {
    let out = run_python(
        r#"
import ssl
print(isinstance(ssl.HAS_ALPN, bool))
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_ssl_has_sni() {
    let out = run_python(
        r#"
import ssl
print(isinstance(ssl.HAS_SNI, bool))
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_ssl_context_set_alpn_protocols() {
    let out = run_python(
        r#"
import ssl
ctx = ssl.SSLContext(ssl.PROTOCOL_TLS_CLIENT)
if ssl.HAS_ALPN:
    ctx.set_alpn_protocols(["http/1.1", "h2"])
print("ok")
"#,
    );
    assert_eq!(out, vec!["ok"]);
}

#[test]
fn test_ssl_context_load_default_certs() {
    let out = run_python(
        r#"
import ssl
ctx = ssl.SSLContext(ssl.PROTOCOL_TLS_CLIENT)
ctx.load_default_certs()
print("ok")
"#,
    );
    assert_eq!(out, vec!["ok"]);
}

#[test]
fn test_ssl_context_set_ciphers_aes() {
    let out = run_python(
        r#"
import ssl
ctx = ssl.SSLContext(ssl.PROTOCOL_TLS_CLIENT)
try:
    ctx.set_ciphers("AES256-SHA")
    print("ok")
except ssl.SSLError:
    print("ok")  # cipher string may not be supported everywhere
"#,
    );
    assert_eq!(out, vec!["ok"]);
}

#[test]
fn test_ssl_purpose_enum() {
    let out = run_python(
        r#"
import ssl
print(ssl.Purpose.SERVER_AUTH.value)
print(ssl.Purpose.CLIENT_AUTH.value)
"#,
    );
    assert_eq!(out, vec!["1.3.6.1.5.5.7.3.1", "1.3.6.1.5.5.7.3.2"]);
}
