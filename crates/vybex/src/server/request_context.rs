//! Build a `vybe_platform_node::http::RequestContext` from an incoming hyper `Request`.
//!
//! Handles CGI-shaped env var population, percent-decoded path, and the
//! response-stream channel that the VM writes into while the async side
//! feeds hyper's body.

use std::collections::HashMap;
use std::net::SocketAddr;
use std::sync::{Arc, Mutex};
use std::time::{SystemTime, UNIX_EPOCH};

use bytes::Bytes;
use http::Request;
use vybe_platform_node::http::{RequestBodyReader, RequestContext, ResponseMessage, ResponseState};

pub struct BuiltContext {
    pub ctx: Arc<RequestContext>,
    pub response_rx: std::sync::mpsc::Receiver<ResponseMessage> }

pub fn build<B>(
    req: &Request<B>,
    body_bytes: Bytes,
    remote: SocketAddr,
    script_filename: Option<&str>,
    script_name: Option<&str>,
    document_root: &std::path::Path,
    local_addr: SocketAddr,
    scheme: &str,
) -> BuiltContext {
    let method = req.method().as_str().to_string();
    let uri_raw = req.uri().to_string();
    let path = percent_encoding::percent_decode_str(req.uri().path())
        .decode_utf8_lossy()
        .into_owned();
    let query = req.uri().query().unwrap_or("").to_string();
    let protocol = format!("{:?}", req.version());
    let host_hdr = req
        .headers()
        .get(http::header::HOST)
        .and_then(|h| h.to_str().ok())
        .unwrap_or("")
        .to_string();

    let (host, port) = split_host_port(&host_hdr, local_addr.port());

    let mut headers = Vec::with_capacity(req.headers().len());
    for (name, value) in req.headers().iter() {
        if let Ok(v) = value.to_str() {
            headers.push((name.as_str().to_string(), v.to_string()));
        }
    }

    let env = build_cgi_env(
        &uri_raw,
        &path,
        &host_hdr,
        &host,
        local_addr,
        port,
        &protocol,
        remote,
        script_filename,
        script_name,
        document_root,
        scheme,
    );

    let body_reader = RequestBodyReader::from_bytes(body_bytes.to_vec());

    let (tx, rx) = std::sync::mpsc::channel::<ResponseMessage>();

    let ctx = RequestContext {
        method,
        uri: uri_raw,
        path,
        query,
        scheme: scheme.to_string(),
        host,
        headers,
        env,
        remote_addr: remote.ip().to_string(),
        remote_port: remote.port(),
        body: Mutex::new(body_reader),
        response: Mutex::new(ResponseState::new(Some(tx))) };

    BuiltContext {
        ctx: Arc::new(ctx),
        response_rx: rx }
}

fn split_host_port(host_hdr: &str, default_port: u16) -> (String, u16) {
    if host_hdr.is_empty() {
        return ("localhost".into(), default_port);
    }
    // Handle IPv6 bracketed form "[::1]:8080".
    if let Some(rest) = host_hdr.strip_prefix('[') {
        if let Some(end) = rest.find(']') {
            let host = &rest[..end];
            let port = rest[end + 1..]
                .strip_prefix(':')
                .and_then(|p| p.parse().ok())
                .unwrap_or(default_port);
            return (host.into(), port);
        }
    }
    match host_hdr.rsplit_once(':') {
        Some((h, p)) => (h.into(), p.parse().unwrap_or(default_port)),
        None => (host_hdr.into(), default_port) }
}

#[allow(clippy::too_many_arguments)]
/// The DEPLOYMENT half of the CGI environment.
///
/// `wasi:http` models the message; it has no document root, resolved script
/// path, server identity or peer address. Those come from the transport, under
/// their standard CGI names so the map stays language-neutral.
///
/// The MESSAGE-derived keys — `REQUEST_METHOD`, `QUERY_STRING`, `REQUEST_URI`,
/// `PATH_INFO`, `CONTENT_TYPE`, `CONTENT_LENGTH`, every `HTTP_*` — are NOT
/// built here. `primitives/http_request_env::emit_environ` derives them from
/// `wasi:http`, so every language gets them including ones that never run
/// under this server. This function used to compute all of them too and then
/// have `publish_server_env` throw them away.
#[allow(clippy::too_many_arguments)]
fn build_cgi_env(
    uri_raw: &str,
    path: &str,
    host_hdr: &str,
    host: &str,
    local_addr: SocketAddr,
    port: u16,
    protocol: &str,
    remote: SocketAddr,
    script_filename: Option<&str>,
    script_name: Option<&str>,
    document_root: &std::path::Path,
    scheme: &str,
) -> HashMap<String, String> {
    let mut env = HashMap::new();
    let now = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .unwrap_or_default();
    let http_host = if host_hdr.is_empty() {
        canonical_http_host(host, port, scheme)
    } else {
        host_hdr.to_string()
    };
    let script_path = script_name.unwrap_or(path);

    env.insert(
        "SERVER_SOFTWARE".into(),
        format!("vybex/{}", env!("CARGO_PKG_VERSION")),
    );
    env.insert("SERVER_NAME".into(), host.into());
    env.insert("SERVER_ADDR".into(), local_addr.ip().to_string());
    env.insert("SERVER_ADMIN".into(), String::new());
    env.insert("SERVER_SIGNATURE".into(), String::new());
    env.insert("SERVER_PORT".into(), port.to_string());
    env.insert("SERVER_PROTOCOL".into(), protocol.into());
    env.insert("GATEWAY_INTERFACE".into(), "CGI/1.1".into());
    env.insert("DOCUMENT_ROOT".into(), document_root.display().to_string());
    env.insert("DOCUMENT_URI".into(), path.into());
    env.insert("REMOTE_HOST".into(), remote.ip().to_string());
    env.insert(
        "SCRIPT_URI".into(),
        absolute_request_uri(uri_raw, scheme, &http_host),
    );
    env.insert("SCRIPT_URL".into(), path.into());
    env.insert("REQUEST_TIME".into(), now.as_secs().to_string());
    env.insert(
        "REQUEST_TIME_FLOAT".into(),
        format!("{}.{:06}", now.as_secs(), now.subsec_micros()),
    );
    if let Some(sf) = script_filename {
        env.insert("SCRIPT_FILENAME".into(), sf.into());
        env.insert("PATH_TRANSLATED".into(), sf.into());
        env.insert("SCRIPT_NAME".into(), script_path.into());
        // `PHP_SELF` is NOT set here. It is PHP's spelling of `SCRIPT_NAME`, and
        // one language's vocabulary in the transport is `php_lang.rs` again —
        // the PHP superglobal prelude derives it.
    }
    env
}

fn canonical_http_host(host: &str, port: u16, scheme: &str) -> String {
    let is_default_port = (scheme == "http" && port == 80) || (scheme == "https" && port == 443);
    if is_default_port {
        host.to_string()
    } else {
        format!("{host}:{port}")
    }
}

fn absolute_request_uri(uri_raw: &str, scheme: &str, http_host: &str) -> String {
    if uri_raw.starts_with("http://") || uri_raw.starts_with("https://") {
        return uri_raw.to_string();
    }
    if uri_raw.starts_with('/') {
        format!("{scheme}://{http_host}{uri_raw}")
    } else {
        format!("{scheme}://{http_host}/{uri_raw}")
    }
}

