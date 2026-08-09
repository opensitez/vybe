//! Bridge the script's `std::sync::mpsc::Receiver<ResponseMessage>` into
//! a hyper streaming body.
//!
//! The script side runs on a blocking worker and writes synchronously via
//! `std::sync::mpsc::Sender`. The hyper side is async. We bridge by
//! forwarding messages onto a bounded `tokio::sync::mpsc::Sender` from a
//! `spawn_blocking` task; a custom `Body` impl pulls from the tokio
//! receiver.
//!
//! The first message MUST be `ResponseMessage::Headers`. We peel it off
//! on the async side and use it to build the hyper response (status +
//! headers). Everything after is body bytes.

use bytes::Bytes;
use http::{HeaderMap, HeaderName, HeaderValue, Response, StatusCode};
use http_body_util::{BodyExt, Full};
use hyper::body::{Body, Frame, SizeHint};
use std::pin::Pin;
use std::sync::mpsc as std_mpsc;
use std::task::{Context, Poll};
use tokio::sync::mpsc as tokio_mpsc;
use vybe_platform_node::http::ResponseMessage;

/// Full boxed hyper body.
pub type BoxBody = http_body_util::combinators::BoxBody<Bytes, std::io::Error>;

/// Body backed by a tokio mpsc receiver. Yields `ResponseMessage::Data`
/// chunks as body frames; ignores stray `Headers` frames after the first.
struct ChannelBody {
    rx: tokio_mpsc::Receiver<ResponseMessage>,
}

impl Body for ChannelBody {
    type Data = Bytes;
    type Error = std::io::Error;

    fn poll_frame(
        mut self: Pin<&mut Self>,
        cx: &mut Context<'_>,
    ) -> Poll<Option<Result<Frame<Self::Data>, Self::Error>>> {
        loop {
            match self.rx.poll_recv(cx) {
                Poll::Pending => return Poll::Pending,
                Poll::Ready(None) => return Poll::Ready(None),
                Poll::Ready(Some(ResponseMessage::Data(bytes))) => {
                    return Poll::Ready(Some(Ok(Frame::data(Bytes::from(bytes)))));
                }
                // Stray headers after first are a script bug; skip and
                // continue polling so we drain to EOF cleanly.
                Poll::Ready(Some(ResponseMessage::Headers { .. })) => continue,
            }
        }
    }

    fn size_hint(&self) -> SizeHint {
        SizeHint::default()
    }
}

/// Await the first `ResponseMessage` (must be `Headers`), then wrap the
/// remaining stream as a hyper body.
///
/// If the script finishes without ever writing a byte (no headers message
/// was ever sent), returns a default 200 empty HTML response.
pub async fn build_response(rx: std_mpsc::Receiver<ResponseMessage>) -> Response<BoxBody> {
    let (tokio_tx, mut tokio_rx) = tokio_mpsc::channel::<ResponseMessage>(16);

    tokio::task::spawn_blocking(move || {
        while let Ok(msg) = rx.recv() {
            if tokio_tx.blocking_send(msg).is_err() {
                break;
            }
        }
    });

    let first = tokio_rx.recv().await;

    let (status, headers) = match first {
        Some(ResponseMessage::Headers { status, headers }) => (status, headers),
        Some(ResponseMessage::Data(_)) => {
            return error_body(500, "internal: body written before headers were flushed");
        }
        None => {
            return default_empty_response();
        }
    };

    let body = ChannelBody { rx: tokio_rx }.boxed();

    let mut resp = Response::builder()
        .status(StatusCode::from_u16(status).unwrap_or(StatusCode::OK))
        .body(body)
        .unwrap();

    apply_headers(resp.headers_mut(), &headers);
    resp
}

/// Build a one-shot (non-streaming) response, for error pages and static
/// files where we have the whole body in memory already.
pub fn bytes_response(status: u16, content_type: &str, body: Vec<u8>) -> Response<BoxBody> {
    let full = Full::new(Bytes::from(body))
        .map_err(|never| match never {})
        .boxed();
    let mut resp = Response::builder()
        .status(StatusCode::from_u16(status).unwrap_or(StatusCode::INTERNAL_SERVER_ERROR))
        .body(full)
        .unwrap();
    resp.headers_mut().insert(
        http::header::CONTENT_TYPE,
        HeaderValue::from_str(content_type)
            .unwrap_or_else(|_| HeaderValue::from_static("application/octet-stream")),
    );
    resp
}

fn default_empty_response() -> Response<BoxBody> {
    let empty = Full::new(Bytes::new())
        .map_err(|never| match never {})
        .boxed();
    let mut resp = Response::builder().status(200).body(empty).unwrap();
    resp.headers_mut().insert(
        http::header::CONTENT_TYPE,
        HeaderValue::from_static("text/html; charset=utf-8"),
    );
    resp
}

fn error_body(status: u16, message: &str) -> Response<BoxBody> {
    bytes_response(
        status,
        "text/plain; charset=utf-8",
        format!("{status} {message}\n").into_bytes(),
    )
}

fn apply_headers(dst: &mut HeaderMap, src: &[(String, String)]) {
    for (n, v) in src {
        let Ok(name) = HeaderName::try_from(n.as_str()) else {
            continue;
        };
        let Ok(value) = HeaderValue::from_str(v) else {
            continue;
        };
        dst.append(name, value);
    }
}
