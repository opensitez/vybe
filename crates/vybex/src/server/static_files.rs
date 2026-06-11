//! Serve static files from disk. Phase 1: simple read-all + Content-Type.
//! Phase 2 will add Range, ETag/If-Modified-Since, gzip negotiation,
//! streaming via tokio::fs::File through a ReaderStream body.

use std::path::Path;

use super::response_stream::{BoxBody, bytes_response};
use http::Response;

pub async fn serve(path: &Path) -> Response<BoxBody> {
    let bytes = match tokio::fs::read(path).await {
        Ok(b) => b,
        Err(_) => {
            return bytes_response(
                500,
                "text/plain; charset=utf-8",
                b"failed to read file\n".to_vec(),
            );
        }
    };

    let mime = mime_guess::from_path(path).first_or_octet_stream();
    let mut content_type = mime.essence_str().to_string();
    // Add charset for textual types.
    if mime.type_() == "text"
        || mime.essence_str() == "application/json"
        || mime.essence_str() == "application/javascript"
        || mime.essence_str() == "application/xml"
        || mime.essence_str() == "image/svg+xml"
    {
        content_type.push_str("; charset=utf-8");
    }

    bytes_response(200, &content_type, bytes)
}
