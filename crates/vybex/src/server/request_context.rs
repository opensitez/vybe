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
    pub response_rx: std::sync::mpsc::Receiver<ResponseMessage>,
}

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
        &method,
        &uri_raw,
        &path,
        &query,
        &host_hdr,
        &host,
        local_addr,
        port,
        &protocol,
        &headers,
        remote,
        script_filename,
        script_name,
        document_root,
        scheme,
        body_bytes.len(),
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
        port,
        protocol,
        headers,
        env,
        remote_addr: remote.ip().to_string(),
        remote_port: remote.port(),
        request_id: new_request_id(req),
        body: Mutex::new(body_reader),
        response: Mutex::new(ResponseState::new(Some(tx))),
        cookies: std::sync::OnceLock::new(),
        query_pairs: std::sync::OnceLock::new(),
    };

    BuiltContext {
        ctx: Arc::new(ctx),
        response_rx: rx,
    }
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
        None => (host_hdr.into(), default_port),
    }
}

#[allow(clippy::too_many_arguments)]
fn build_cgi_env(
    method: &str,
    uri_raw: &str,
    path: &str,
    query: &str,
    host_hdr: &str,
    host: &str,
    local_addr: SocketAddr,
    port: u16,
    protocol: &str,
    headers: &[(String, String)],
    remote: SocketAddr,
    script_filename: Option<&str>,
    script_name: Option<&str>,
    document_root: &std::path::Path,
    scheme: &str,
    body_len: usize,
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
    let script_uri = absolute_request_uri(uri_raw, scheme, &http_host);

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
    env.insert("REQUEST_METHOD".into(), method.into());
    env.insert("REQUEST_SCHEME".into(), scheme.into());
    env.insert("REQUEST_URI".into(), uri_raw.into());
    env.insert("DOCUMENT_URI".into(), path.into());
    env.insert("PATH_INFO".into(), path.into());
    env.insert("QUERY_STRING".into(), query.into());
    env.insert("DOCUMENT_ROOT".into(), document_root.display().to_string());
    env.insert("REMOTE_ADDR".into(), remote.ip().to_string());
    env.insert("REMOTE_HOST".into(), remote.ip().to_string());
    env.insert("REMOTE_PORT".into(), remote.port().to_string());
    env.insert("HTTP_HOST".into(), http_host.clone());
    env.insert("SCRIPT_URI".into(), script_uri);
    env.insert("SCRIPT_URL".into(), path.into());
    env.insert("REQUEST_TIME".into(), now.as_secs().to_string());
    env.insert(
        "REQUEST_TIME_FLOAT".into(),
        format!("{}.{:06}", now.as_secs(), now.subsec_micros()),
    );
    if scheme == "https" {
        env.insert("HTTPS".into(), "on".into());
    }
    if let Some(sf) = script_filename {
        env.insert("SCRIPT_FILENAME".into(), sf.into());
        env.insert("PATH_TRANSLATED".into(), sf.into());
        env.insert("SCRIPT_NAME".into(), script_path.into());
        env.insert("PHP_SELF".into(), script_path.into());
    }
    if body_len > 0 {
        env.insert("CONTENT_LENGTH".into(), body_len.to_string());
    }

    // Collect per-name values so duplicates are joined with ", " per RFC 3875.
    let mut by_key: HashMap<String, Vec<&str>> = HashMap::new();
    for (n, v) in headers {
        let key = n.as_str();
        if key.eq_ignore_ascii_case("content-type") {
            env.entry("CONTENT_TYPE".into())
                .or_insert_with(|| v.clone());
            continue;
        }
        if key.eq_ignore_ascii_case("content-length") {
            continue; // already set above from body_len
        }
        let http_name = format!("HTTP_{}", key.to_ascii_uppercase().replace('-', "_"));
        by_key.entry(http_name).or_default().push(v.as_str());
    }
    for (k, vs) in by_key {
        env.insert(k, vs.join(", "));
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

fn new_request_id<B>(req: &Request<B>) -> String {
    if let Some(v) = req
        .headers()
        .get("x-request-id")
        .and_then(|v| v.to_str().ok())
    {
        // Honor inbound if it's a sane length + charset.
        if (8..=128).contains(&v.len())
            && v.chars()
                .all(|c| c.is_ascii_alphanumeric() || c == '-' || c == '_')
        {
            return v.to_string();
        }
    }
    uuid::Uuid::new_v4().to_string()
}
