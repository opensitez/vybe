//! Dev error pages. Phase 1: minimal HTML; Phase 2+ gets source snippets
//! with highlighted line numbers.

use super::response_stream::{bytes_response, BoxBody};
use http::Response;

pub fn error_404(path: &str) -> Response<BoxBody> {
    let body = format!(
        "<!doctype html><html><head><title>404 Not Found</title></head>\
         <body style=\"font-family:system-ui;margin:2em\">\
         <h1>404 Not Found</h1><p><code>{}</code> does not exist on this server.</p>\
         <hr><small>vybex {}</small></body></html>",
        html_escape(path),
        env!("CARGO_PKG_VERSION"),
    );
    bytes_response(404, "text/html; charset=utf-8", body.into_bytes())
}

pub fn error_403(path: &str) -> Response<BoxBody> {
    let body = format!(
        "<!doctype html><html><head><title>403 Forbidden</title></head>\
         <body style=\"font-family:system-ui;margin:2em\">\
         <h1>403 Forbidden</h1><p><code>{}</code> is not accessible.</p>\
         <hr><small>vybex {}</small></body></html>",
        html_escape(path),
        env!("CARGO_PKG_VERSION"),
    );
    bytes_response(403, "text/html; charset=utf-8", body.into_bytes())
}

pub fn error_501(msg: &str) -> Response<BoxBody> {
    let body = format!(
        "<!doctype html><html><head><title>501 Not Implemented</title></head>\
         <body style=\"font-family:system-ui;margin:2em\">\
         <h1>501 Not Implemented</h1><p>{}</p>\
         <hr><small>vybex {}</small></body></html>",
        html_escape(msg),
        env!("CARGO_PKG_VERSION"),
    );
    bytes_response(501, "text/html; charset=utf-8", body.into_bytes())
}

fn html_escape(s: &str) -> String {
    s.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
}
