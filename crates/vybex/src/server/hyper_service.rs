//! Tower `Service` that hyper invokes per incoming HTTP request.
//!
//! Responsibilities:
//! - Reject disallowed methods (Phase 1 accepts GET/HEAD/POST/PUT/DELETE/PATCH/OPTIONS).
//! - Enforce `max_body` via `tower_http::limit::RequestBodyLimitLayer`
//!   (wired in `mod.rs`).
//! - Resolve the URL to a filesystem path via `directory::resolve`.
//! - Dispatch to `static_files::serve` or `script::serve` based on extension.

use std::convert::Infallible;
use std::net::SocketAddr;
use std::sync::Arc;

use bytes::Bytes;
use http::{Method, Request, Response};
use http_body_util::BodyExt;
use hyper::body::Incoming;

use super::config::ServeConfig;
use super::response_stream::BoxBody;

/// Construct the tower service hyper will call for each request on a
/// given connection. We close over `config` and the `remote` addr.
pub fn make_service(config: Arc<ServeConfig>, remote: SocketAddr) -> ServeHandler {
    ServeHandler { config, remote }
}

#[derive(Clone)]
pub struct ServeHandler {
    config: Arc<ServeConfig>,
    remote: SocketAddr,
}

impl hyper::service::Service<Request<Incoming>> for ServeHandler {
    type Response = Response<BoxBody>;
    type Error = Infallible;
    type Future = std::pin::Pin<
        Box<dyn std::future::Future<Output = Result<Response<BoxBody>, Infallible>> + Send>,
    >;

    fn call(&self, req: Request<Incoming>) -> Self::Future {
        let config = Arc::clone(&self.config);
        let remote = self.remote;
        Box::pin(async move {
            let start = std::time::Instant::now();
            let method = req.method().clone();
            let path = req.uri().path().to_string();
            let protocol = format!("{:?}", req.version());

            let response = handle(req, config, remote).await;

            super::logging::log_request(
                remote,
                method.as_str(),
                &path,
                &protocol,
                response.status().as_u16(),
                start.elapsed(),
            );
            Ok(response)
        })
    }
}

async fn handle(
    req: Request<Incoming>,
    config: Arc<ServeConfig>,
    remote: SocketAddr,
) -> Response<BoxBody> {
    // Method gate: Phase 1 accepts common safe methods. Unsupported →
    // 501 Not Implemented.
    match req.method() {
        &Method::GET
        | &Method::HEAD
        | &Method::POST
        | &Method::PUT
        | &Method::DELETE
        | &Method::PATCH
        | &Method::OPTIONS => {}
        other => {
            return super::errors::error_501(&format!("method {other} not supported"));
        }
    }

    // Resolve URL path → filesystem path.
    let url_path = req.uri().path().to_string();
    let resolution = super::directory::resolve(&url_path, &config);

    let path = match resolution {
        super::directory::Resolution::File(p) => p,
        super::directory::Resolution::NotFound => return super::errors::error_404(&url_path),
        super::directory::Resolution::Forbidden => return super::errors::error_403(&url_path),
    };

    // Static files: fast path, no VM involvement.
    if !super::directory::is_script(&path) {
        return super::static_files::serve(&path).await;
    }

    // Script path: collect body, build RequestContext, dispatch to VM.
    let (parts, body) = req.into_parts();
    let body_bytes = match read_limited(body, config.max_body).await {
        Ok(b) => b,
        Err(BodyError::TooLarge) => {
            return super::errors::error_501("request body exceeds --max-body");
        }
        Err(BodyError::Read(e)) => {
            return super::errors::error_501(&format!("body read error: {e}"));
        }
    };

    let local_addr: SocketAddr = config
        .bind
        .parse()
        .unwrap_or_else(|_| "127.0.0.1:8080".parse().unwrap());

    let scheme = "http"; // Phase 2: "https" when TLS is wired.
    let reassembled = Request::from_parts(parts, ());
    let built = super::request_context::build(
        &reassembled,
        body_bytes,
        remote,
        Some(path.to_string_lossy().as_ref()),
        path.strip_prefix(&config.root)
            .ok()
            .map(|rel| format!("/{}", rel.to_string_lossy().replace('\\', "/")))
            .as_deref(),
        &config.root,
        local_addr,
        scheme,
    );

    super::script::serve(
        path,
        built.ctx,
        built.response_rx,
        config.no_sandbox,
        config.timeout_secs,
        config.shutdown.clone(),
    )
    .await
}

#[derive(Debug)]
enum BodyError {
    TooLarge,
    Read(hyper::Error),
}

async fn read_limited(mut body: Incoming, max: usize) -> Result<Bytes, BodyError> {
    let mut buf = Vec::<u8>::new();
    while let Some(frame) = body.frame().await {
        let frame = frame.map_err(BodyError::Read)?;
        if let Some(data) = frame.data_ref() {
            if buf.len() + data.len() > max {
                return Err(BodyError::TooLarge);
            }
            buf.extend_from_slice(data);
        }
    }
    Ok(Bytes::from(buf))
}
