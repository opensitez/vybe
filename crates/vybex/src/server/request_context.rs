//! Build a `vybe_host::RequestContext` from an incoming hyper `Request`.
//!
//! Handles CGI-shaped env var population, percent-decoded path, and the
//! response-stream channel that the VM writes into while the async side
//! feeds hyper's body.

use std::collections::HashMap;
use std::net::SocketAddr;
use std::sync::{Arc, Mutex};

use bytes::Bytes;
use http::Request;
use vybe_host::{RequestBodyReader, RequestContext, ResponseMessage, ResponseState};

pub struct BuiltContext {
    pub ctx: Arc<RequestContext>,
    pub response_rx: std::sync::mpsc::Receiver<ResponseMessage>,
}

pub fn build<B>(
    req: &Request<B>,
    body_bytes: Bytes,
    remote: SocketAddr,
    script_filename: Option<&str>,
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
        &method, &uri_raw, &path, &query, &host, port,
        &protocol, &headers, remote, script_filename, document_root, scheme,
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
    };

    BuiltContext { ctx: Arc::new(ctx), response_rx: rx }
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
    host: &str,
    port: u16,
    protocol: &str,
    headers: &[(String, String)],
    remote: SocketAddr,
    script_filename: Option<&str>,
    document_root: &std::path::Path,
    scheme: &str,
    body_len: usize,
) -> HashMap<String, String> {
    let mut env = HashMap::new();
    env.insert("SERVER_SOFTWARE".into(), format!("vybex/{}", env!("CARGO_PKG_VERSION")));
    env.insert("SERVER_NAME".into(), host.into());
    env.insert("SERVER_PORT".into(), port.to_string());
    env.insert("SERVER_PROTOCOL".into(), protocol.into());
    env.insert("GATEWAY_INTERFACE".into(), "CGI/1.1".into());
    env.insert("REQUEST_METHOD".into(), method.into());
    env.insert("REQUEST_URI".into(), uri_raw.into());
    env.insert("PATH_INFO".into(), path.into());
    env.insert("QUERY_STRING".into(), query.into());
    env.insert("DOCUMENT_ROOT".into(), document_root.display().to_string());
    env.insert("REMOTE_ADDR".into(), remote.ip().to_string());
    env.insert("REMOTE_PORT".into(), remote.port().to_string());
    if scheme == "https" {
        env.insert("HTTPS".into(), "on".into());
    }
    if let Some(sf) = script_filename {
        env.insert("SCRIPT_FILENAME".into(), sf.into());
        env.insert("SCRIPT_NAME".into(), path.into());
    }
    if body_len > 0 {
        env.insert("CONTENT_LENGTH".into(), body_len.to_string());
    }

    // Collect per-name values so duplicates are joined with ", " per RFC 3875.
    let mut by_key: HashMap<String, Vec<&str>> = HashMap::new();
    for (n, v) in headers {
        let key = n.as_str();
        if key.eq_ignore_ascii_case("content-type") {
            env.entry("CONTENT_TYPE".into()).or_insert_with(|| v.clone());
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

fn new_request_id<B>(req: &Request<B>) -> String {
    if let Some(v) = req.headers().get("x-request-id").and_then(|v| v.to_str().ok()) {
        // Honor inbound if it's a sane length + charset.
        if (8..=128).contains(&v.len())
            && v.chars().all(|c| c.is_ascii_alphanumeric() || c == '-' || c == '_')
        {
            return v.to_string();
        }
    }
    uuid::Uuid::new_v4().to_string()
}
