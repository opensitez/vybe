//! Per-request state for `vybe:http`.
//!
//! One `RequestContext` is constructed per incoming HTTP request, installed
//! as a thread-local on the blocking worker that runs the script, and
//! drained by the server to form the HTTP response. A `ContextGuard` with a
//! `Drop` implementation ensures the thread-local is always cleared,
//! including on panic.

use std::cell::RefCell;
use std::collections::HashMap;
use std::io::Cursor;
use std::sync::{Arc, Mutex};

/// Streaming reader over the request body.
///
/// Phase 1: backed by an in-memory `Cursor<Vec<u8>>` (body collected up to
/// `max-body` before the VM runs). Phase 2 swaps in a reader that pulls
/// from the async hyper body via a bridge channel.
pub struct RequestBodyReader {
    inner: Cursor<Vec<u8>>,
    length: Option<usize>,
}

impl RequestBodyReader {
    pub fn from_bytes(bytes: Vec<u8>) -> Self {
        let length = Some(bytes.len());
        Self { inner: Cursor::new(bytes), length }
    }

    pub fn empty() -> Self {
        Self { inner: Cursor::new(Vec::new()), length: Some(0) }
    }

    pub fn length(&self) -> Option<usize> { self.length }

    pub fn eof(&self) -> bool {
        self.inner.position() as usize >= self.inner.get_ref().len()
    }

    pub fn read(&mut self, max: usize) -> Vec<u8> {
        use std::io::Read;
        let mut buf = vec![0u8; max];
        match self.inner.read(&mut buf) {
            Ok(n) => { buf.truncate(n); buf }
            Err(_) => Vec::new(),
        }
    }

    pub fn read_all(&mut self) -> Vec<u8> {
        use std::io::Read;
        let mut out = Vec::new();
        let _ = self.inner.read_to_end(&mut out);
        out
    }
}

/// One outbound message from the script thread to the hyper writer task.
///
/// The first message on a response stream must be `Headers`. Subsequent
/// messages are `Data` chunks until the sender is dropped or the script
/// calls `response.end()`, at which point the hyper body stream closes.
#[derive(Debug)]
pub enum ResponseMessage {
    Headers { status: u16, headers: Vec<(String, String)> },
    Data(Vec<u8>),
}

/// Per-request response state owned by the script side.
///
/// The server side holds the matching `mpsc::Receiver<ResponseMessage>`
/// which it drains into a hyper streaming body. `headers_sent` flips true
/// on the first `write` or explicit flush; further header mutations after
/// that are ignored (with a debug warning).
pub struct ResponseState {
    pub status: u16,
    pub headers: Vec<(String, String)>,
    pub headers_sent: bool,
    pub ended: bool,
    /// Channel to the hyper writer task. `None` in "cli" mode (no server).
    pub sender: Option<std::sync::mpsc::Sender<ResponseMessage>>,
}

impl ResponseState {
    pub fn new(sender: Option<std::sync::mpsc::Sender<ResponseMessage>>) -> Self {
        Self {
            status: 200,
            headers: Vec::new(),
            headers_sent: false,
            ended: false,
            sender,
        }
    }

    /// Flush status + headers on first body write. Idempotent.
    fn flush_headers(&mut self) {
        if self.headers_sent { return; }
        if let Some(tx) = &self.sender {
            let _ = tx.send(ResponseMessage::Headers {
                status: self.status,
                headers: std::mem::take(&mut self.headers),
            });
        }
        self.headers_sent = true;
    }

    pub fn write_bytes(&mut self, bytes: Vec<u8>) {
        if self.ended { return; }
        self.flush_headers();
        if let Some(tx) = &self.sender {
            let _ = tx.send(ResponseMessage::Data(bytes));
        }
    }

    pub fn end(&mut self) {
        if self.ended { return; }
        self.flush_headers();
        self.ended = true;
        self.sender = None; // drop sender; hyper sees body EOF
    }
}

/// Per-request state owned by a single blocking worker thread.
///
/// The struct is wrapped in `Arc` and installed into the thread-local
/// `REQUEST_CONTEXT` before the script runs. Host functions read it via
/// [`with_context`]. The `ContextGuard` restores the previous value on
/// drop, so panics and early returns do not leak state into the next
/// request the worker handles.
pub struct RequestContext {
    // Raw, populated at construction.
    pub method: String,
    pub uri: String,
    pub path: String,
    pub query: String,
    pub scheme: String,
    pub host: String,
    pub port: u16,
    pub protocol: String,
    pub headers: Vec<(String, String)>,
    pub env: HashMap<String, String>,
    pub remote_addr: String,
    pub remote_port: u16,
    pub request_id: String,

    // Streaming body reader.
    pub body: Mutex<RequestBodyReader>,

    // Outbound response state.
    pub response: Mutex<ResponseState>,
}

thread_local! {
    /// Currently-executing request's context. `None` in CLI mode.
    static REQUEST_CONTEXT: RefCell<Option<Arc<RequestContext>>> = const { RefCell::new(None) };
}

/// Install `ctx` as the active request context on the current thread and
/// return a guard that clears it on drop (including panic unwind).
///
/// The guard captures the *previous* value so nested installs are valid
/// (Phase 5+ WebSocket upgrades may need this).
#[must_use = "dropping the guard immediately clears the context"]
pub fn install_context(ctx: Arc<RequestContext>) -> ContextGuard {
    let prev = REQUEST_CONTEXT.with(|slot| slot.replace(Some(ctx)));
    ContextGuard { prev }
}

/// Restore-on-drop guard returned by [`install_context`].
pub struct ContextGuard {
    prev: Option<Arc<RequestContext>>,
}

impl Drop for ContextGuard {
    fn drop(&mut self) {
        let prev = self.prev.take();
        REQUEST_CONTEXT.with(|slot| *slot.borrow_mut() = prev);
    }
}

/// Take (clear) the context on the current thread, returning it. Used by
/// the server to recover the response state after the VM finishes.
pub fn take_context() -> Option<Arc<RequestContext>> {
    REQUEST_CONTEXT.with(|slot| slot.borrow_mut().take())
}

/// Run `f` with a shared reference to the current request context.
/// Returns `None` when called in CLI mode (no server running).
pub fn with_context<R>(f: impl FnOnce(&RequestContext) -> R) -> Option<R> {
    REQUEST_CONTEXT.with(|slot| slot.borrow().as_ref().map(|arc| f(arc)))
}
