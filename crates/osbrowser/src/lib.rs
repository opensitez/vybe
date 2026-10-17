//! A browser-owned DOM connected to a native Vybe process.
//!
//! The page runs in the system browser. Vybe retains its VM and sends typed
//! operations over a loopback HTTP channel; DOM events come back on the same
//! channel. No browser executable, UI toolkit, or rendering engine is linked.

use std::collections::{HashMap, VecDeque};
use std::net::SocketAddr;
use std::process::Command;
use std::sync::{Arc, Condvar, Mutex, mpsc};
use std::time::{Duration, Instant};

use bytes::Bytes;
use http_body_util::{BodyExt, Full};
use hyper::body::Incoming;
use hyper::server::conn::http1;
use hyper::service::service_fn;
use hyper::{Method, Request, Response, StatusCode};
use hyper_util::rt::TokioIo;
use serde::{Deserialize, Serialize};
use serde_json::Value;

#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct BrowserEvent {
    pub document: u64,
    pub node: u64,
    #[serde(default)]
    pub path: Vec<u64>,
    pub kind: String,
    #[serde(default)]
    pub fields: Value,
}

#[derive(Serialize)]
struct WireCommand {
    id: u64,
    document: u64,
    operation: String,
    args: Value,
    #[serde(skip_serializing_if = "Option::is_none")]
    assigned_node: Option<u64>,
    #[serde(skip)]
    pixels: Option<Bytes>,
    #[serde(skip_serializing_if = "Option::is_none")]
    binary_range: Option<[usize; 2]>,
}

fn command_response(mut commands: Vec<WireCommand>) -> Response<Body> {
    if !commands.iter().any(|command| command.pixels.is_some()) {
        return response(
            StatusCode::OK,
            serde_json::to_vec(&commands).unwrap_or_default(),
            "application/json",
        );
    }
    let mut offset = 0;
    for command in &mut commands {
        if let Some(pixels) = &command.pixels {
            command.binary_range = Some([offset, pixels.len()]);
            offset += pixels.len();
        }
    }
    let header = serde_json::to_vec(&commands).expect("serializable commands");
    let Ok(header_len) = u32::try_from(header.len()) else {
        return response(
            StatusCode::INTERNAL_SERVER_ERROR,
            "command header too large",
            "text/plain",
        );
    };
    let mut body = Vec::with_capacity(4 + header.len() + offset);
    body.extend_from_slice(&header_len.to_le_bytes());
    body.extend_from_slice(&header);
    for command in commands {
        if let Some(pixels) = command.pixels {
            body.extend_from_slice(&pixels);
        }
    }
    response(StatusCode::OK, body, "application/octet-stream")
}

#[derive(Deserialize)]
struct WireReply {
    id: u64,
    result: Value,
    #[serde(default)]
    error: Option<String>,
}

#[derive(Deserialize)]
#[serde(untagged)]
enum EventPacket {
    Batch(Vec<BrowserEvent>),
    Single(BrowserEvent),
}

struct State {
    next: u64,
    commands: VecDeque<WireCommand>,
    replies: HashMap<u64, mpsc::Sender<Result<Value, String>>>,
    events: VecDeque<BrowserEvent>,
    connected: bool,
    last_seen: Option<Instant>,
    batch_depth: usize,
    flush_requested: bool,
}

impl Default for State {
    fn default() -> Self {
        Self {
            next: 1,
            commands: VecDeque::new(),
            replies: HashMap::new(),
            events: VecDeque::new(),
            connected: false,
            last_seen: None,
            batch_depth: 0,
            flush_requested: false,
        }
    }
}

impl State {
    fn take_ready_commands(&mut self) -> Vec<WireCommand> {
        self.last_seen = Some(Instant::now());
        if self.batch_depth > 0 && !self.flush_requested {
            return Vec::new();
        }
        self.flush_requested = false;
        self.commands.drain(..).collect()
    }
}

/// One browser tab and its bidirectional command channel.
pub struct BrowserSession {
    address: SocketAddr,
    token: String,
    state: Arc<Mutex<State>>,
    event_ready: Arc<Condvar>,
    wake: Arc<tokio::sync::Notify>,
}

impl BrowserSession {
    /// Start a loopback server. The browser is opened separately with `open`.
    pub fn start() -> Result<Self, String> {
        let (ready_tx, ready_rx) = mpsc::sync_channel(1);
        let state = Arc::new(Mutex::new(State::default()));
        let server_state = Arc::clone(&state);
        let event_ready = Arc::new(Condvar::new());
        let server_event_ready = Arc::clone(&event_ready);
        let token = uuid::Uuid::new_v4().simple().to_string();
        let server_token = token.clone();
        let wake = Arc::new(tokio::sync::Notify::new());
        let server_wake = Arc::clone(&wake);
        std::thread::Builder::new()
            .name("osbrowser-http".into())
            .spawn(move || {
                let runtime = tokio::runtime::Builder::new_current_thread()
                    .enable_all()
                    .build();
                let Ok(runtime) = runtime else {
                    let _ = ready_tx.send(Err("cannot start browser runtime".to_string()));
                    return;
                };
                runtime.block_on(async move {
                    let listener = tokio::net::TcpListener::bind("127.0.0.1:0").await;
                    let Ok(listener) = listener else {
                        let _ = ready_tx.send(Err("cannot bind browser bridge".to_string()));
                        return;
                    };
                    let _ = ready_tx.send(listener.local_addr().map_err(|e| e.to_string()));
                    loop {
                        let Ok((stream, _)) = listener.accept().await else {
                            break;
                        };
                        let state = Arc::clone(&server_state);
                        let event_ready = Arc::clone(&server_event_ready);
                        let token = server_token.clone();
                        let wake = Arc::clone(&server_wake);
                        tokio::spawn(async move {
                            let service = service_fn(move |request| {
                                serve(
                                    request,
                                    Arc::clone(&state),
                                    token.clone(),
                                    Arc::clone(&wake),
                                    Arc::clone(&event_ready),
                                )
                            });
                            let _ = http1::Builder::new()
                                .serve_connection(TokioIo::new(stream), service)
                                .await;
                        });
                    }
                });
            })
            .map_err(|e| e.to_string())?;
        Ok(Self {
            address: ready_rx.recv().map_err(|e| e.to_string())??,
            token,
            state,
            event_ready,
            wake,
        })
    }

    pub fn url(&self) -> String {
        format!("http://{}/{}/", self.address, self.token)
    }

    /// Open an external browser window. Prefer Chrome on macOS, where the
    /// bridge is exercised against Chrome; fall back to the system browser.
    pub fn open(&self) -> Result<(), String> {
        let url = self.url();
        #[cfg(target_os = "macos")]
        let status = if std::path::Path::new("/Applications/Google Chrome.app").exists() {
            Command::new("open")
                .args(["-a", "Google Chrome", &url])
                .status()
        } else {
            Command::new("open").arg(&url).status()
        };
        #[cfg(target_os = "windows")]
        let status = Command::new("cmd").args(["/C", "start", "", &url]).status();
        #[cfg(all(unix, not(target_os = "macos")))]
        let status = Command::new("xdg-open").arg(&url).status();
        #[cfg(not(any(unix, target_os = "windows")))]
        let status: std::io::Result<std::process::ExitStatus> =
            Err(std::io::Error::other("unsupported platform"));
        match status {
            Ok(code) if code.success() => Ok(()),
            Ok(code) => Err(format!("browser launcher exited with {code}")),
            Err(error) => Err(error.to_string()),
        }
    }

    pub fn is_connected(&self) -> bool {
        self.state.lock().is_ok_and(|state| {
            state.connected
                && state
                    .last_seen
                    .is_some_and(|at| at.elapsed() < Duration::from_secs(10))
        })
    }

    pub fn begin_batch(&self) {
        if let Ok(mut state) = self.state.lock() {
            state.batch_depth += 1;
        }
    }

    pub fn end_batch(&self) {
        if let Ok(mut state) = self.state.lock() {
            state.batch_depth = state.batch_depth.saturating_sub(1);
            if state.batch_depth == 0 && !state.commands.is_empty() {
                drop(state);
                self.wake.notify_one();
            }
        }
    }

    /// Send a DOM operation and wait for the browser's reply.
    pub fn call(&self, document: u64, operation: &str, args: Value) -> Result<Value, String> {
        self.call_payload(document, operation, args, None)
    }

    /// Upload raw RGBA bytes using standard Canvas ImageData metadata.
    pub fn call_with_pixels(
        &self,
        document: u64,
        args: Value,
        pixels: Bytes,
    ) -> Result<Value, String> {
        self.call_payload(document, "CanvasApply", args, Some(pixels))
    }

    fn call_payload(
        &self,
        document: u64,
        operation: &str,
        args: Value,
        pixels: Option<Bytes>,
    ) -> Result<Value, String> {
        let (tx, rx) = mpsc::channel();
        let id = {
            let mut state = self.state.lock().map_err(|e| e.to_string())?;
            let id = state.next;
            state.next += 1;
            state.replies.insert(id, tx);
            state.commands.push_back(WireCommand {
                id,
                document,
                operation: operation.to_string(),
                args,
                assigned_node: None,
                pixels,
                binary_range: None,
            });
            state.flush_requested = true;
            id
        };
        self.wake.notify_one();
        let result = rx.recv_timeout(Duration::from_secs(10));
        if result.is_err() {
            if let Ok(mut state) = self.state.lock() {
                state.replies.remove(&id);
                state.commands.retain(|command| command.id != id);
            }
        }
        result.map_err(|_| format!("browser did not answer {operation}"))?
    }

    /// Queue an ordered browser mutation without waiting for a round trip.
    /// The next `call` observes all preceding queued operations.
    pub fn enqueue(
        &self,
        document: u64,
        operation: &str,
        args: Value,
        assigned_node: Option<u64>,
    ) -> Result<(), String> {
        self.enqueue_payload(document, operation, args, assigned_node, None)
    }

    /// Queue a raw RGBA upload in the same ordered stream as DOM mutations.
    pub fn enqueue_with_pixels(
        &self,
        document: u64,
        args: Value,
        pixels: Bytes,
    ) -> Result<(), String> {
        self.enqueue_payload(document, "CanvasApply", args, None, Some(pixels))
    }

    fn enqueue_payload(
        &self,
        document: u64,
        operation: &str,
        args: Value,
        assigned_node: Option<u64>,
        pixels: Option<Bytes>,
    ) -> Result<(), String> {
        let mut state = self.state.lock().map_err(|e| e.to_string())?;
        state.commands.push_back(WireCommand {
            id: 0,
            document,
            operation: operation.to_string(),
            args,
            assigned_node,
            pixels,
            binary_range: None,
        });
        let ready = state.batch_depth == 0;
        drop(state);
        if ready {
            self.wake.notify_one();
        }
        Ok(())
    }

    pub fn drain_events(&self) -> Vec<BrowserEvent> {
        self.state
            .lock()
            .map(|mut state| state.events.drain(..).collect())
            .unwrap_or_default()
    }

    pub fn wait_for_event(&self, timeout: Duration) {
        if let Ok(state) = self.state.lock() {
            if state.events.is_empty() {
                let _ = self
                    .event_ready
                    .wait_timeout_while(state, timeout, |state| state.events.is_empty());
            }
        }
    }
}

type Body = Full<Bytes>;

fn response(
    status: StatusCode,
    body: impl Into<Bytes>,
    content_type: &'static str,
) -> Response<Body> {
    Response::builder()
        .status(status)
        .header("content-type", content_type)
        .header("cache-control", "no-store")
        .body(Full::new(body.into()))
        .expect("fixed response headers")
}

async fn serve(
    request: Request<Incoming>,
    state: Arc<Mutex<State>>,
    token: String,
    wake: Arc<tokio::sync::Notify>,
    event_ready: Arc<Condvar>,
) -> Result<Response<Body>, std::convert::Infallible> {
    let prefix = format!("/{token}");
    let path = request.uri().path().strip_prefix(&prefix).unwrap_or("");
    let answer = match (request.method(), path) {
        (&Method::GET, "/") => response(
            StatusCode::OK,
            include_str!("bridge.html"),
            "text/html; charset=utf-8",
        ),
        (&Method::GET, "/bridge.js") => response(
            StatusCode::OK,
            include_str!("bridge.js"),
            "text/javascript; charset=utf-8",
        ),
        (&Method::GET, "/canvas.js") => response(
            StatusCode::OK,
            include_str!("canvas.js"),
            "text/javascript; charset=utf-8",
        ),
        (&Method::GET, "/next") => {
            let mut commands = state
                .lock()
                .ok()
                .map(|mut state| {
                    state.connected = true;
                    state.take_ready_commands()
                })
                .unwrap_or_default();
            if commands.is_empty() {
                let _ = tokio::time::timeout(Duration::from_secs(1), wake.notified()).await;
                commands = state
                    .lock()
                    .ok()
                    .map(|mut state| state.take_ready_commands())
                    .unwrap_or_default();
            }
            if commands.is_empty() {
                response(StatusCode::NO_CONTENT, Bytes::new(), "application/json")
            } else {
                command_response(commands)
            }
        }
        (&Method::POST, "/reply") => {
            let body = request.into_body().collect().await;
            if let Ok(body) = body {
                if let Ok(replies) = serde_json::from_slice::<Vec<WireReply>>(&body.to_bytes()) {
                    if let Ok(mut state) = state.lock() {
                        for reply in replies {
                            if let Some(tx) = state.replies.remove(&reply.id) {
                                let _ = tx.send(match reply.error {
                                    Some(error) => Err(error),
                                    None => Ok(reply.result),
                                });
                            }
                        }
                    }
                }
            }
            response(StatusCode::NO_CONTENT, Bytes::new(), "application/json")
        }
        (&Method::POST, "/event") => {
            let body = request.into_body().collect().await;
            if let Ok(body) = body {
                if let Ok(packet) = serde_json::from_slice::<EventPacket>(&body.to_bytes()) {
                    if let Ok(mut state) = state.lock() {
                        match packet {
                            EventPacket::Batch(events) => state.events.extend(events),
                            EventPacket::Single(event) => state.events.push_back(event),
                        }
                        event_ready.notify_one();
                    }
                }
            }
            response(StatusCode::NO_CONTENT, Bytes::new(), "application/json")
        }
        _ => response(StatusCode::NOT_FOUND, "not found", "text/plain"),
    };
    Ok(answer)
}

#[cfg(test)]
mod tests {
    use super::{BrowserEvent, BrowserSession, State, WireCommand};
    use std::sync::{Arc, Condvar, Mutex};
    use std::time::{Duration, Instant};

    fn command(id: u64, operation: &str) -> WireCommand {
        WireCommand {
            id,
            document: 1,
            operation: operation.into(),
            args: serde_json::Value::Null,
            assigned_node: None,
            pixels: None,
            binary_range: None,
        }
    }

    #[test]
    fn event_packets_accept_legacy_and_ordered_batches() {
        let down = serde_json::json!({"document": 1, "node": 2,
            "kind": "keydown", "fields": {"key": "ArrowUp"}});
        let up = serde_json::json!({"document": 1, "node": 2,
            "kind": "keyup", "fields": {"key": "ArrowUp"}});
        let single: super::EventPacket = serde_json::from_value(down.clone()).unwrap();
        assert!(matches!(single, super::EventPacket::Single(event) if event.kind == "keydown"));
        let batch: super::EventPacket =
            serde_json::from_value(serde_json::json!([down, up])).unwrap();
        let super::EventPacket::Batch(events) = batch else {
            panic!("expected batch")
        };
        assert_eq!(
            events
                .iter()
                .map(|event| event.kind.as_str())
                .collect::<Vec<_>>(),
            ["keydown", "keyup"]
        );
    }

    #[test]
    fn binary_response_preserves_order_and_raw_pixel_ranges() {
        use http_body_util::BodyExt;
        let mut first = command(1, "CanvasApply");
        first.pixels = Some(bytes::Bytes::from_static(&[1, 2, 3, 255]));
        let second = command(2, "Title");
        let mut third = command(3, "CanvasApply");
        third.pixels = Some(bytes::Bytes::from_static(&[4, 5, 6, 255]));
        let response = super::command_response(vec![first, second, third]);
        assert_eq!(
            response.headers()["content-type"],
            "application/octet-stream"
        );
        let runtime = tokio::runtime::Builder::new_current_thread()
            .build()
            .unwrap();
        let body = runtime
            .block_on(response.into_body().collect())
            .unwrap()
            .to_bytes();
        let length = u32::from_le_bytes(body[..4].try_into().unwrap()) as usize;
        let commands: serde_json::Value = serde_json::from_slice(&body[4..4 + length]).unwrap();
        assert_eq!(commands[0]["binary_range"], serde_json::json!([0, 4]));
        assert!(commands[1].get("binary_range").is_none());
        assert_eq!(commands[2]["binary_range"], serde_json::json!([4, 4]));
        assert_eq!(&body[4 + length..], &[1, 2, 3, 255, 4, 5, 6, 255]);
    }

    #[test]
    fn mutations_wait_for_batch_end() {
        let mut state = State::default();
        state.batch_depth = 1;
        state.commands.push_back(command(0, "SetStyleProperty"));
        assert!(state.take_ready_commands().is_empty());
        state.batch_depth = 0;
        assert_eq!(state.take_ready_commands()[0].operation, "SetStyleProperty");
    }

    #[test]
    fn grid_creation_drains_as_one_ordered_batch() {
        let mut state = State::default();
        state.batch_depth = 1;
        for _ in 0..42 {
            state.commands.push_back(command(0, "CreateElement"));
            state.commands.push_back(command(0, "SetStyleProperty"));
            state.commands.push_back(command(0, "SetStyleProperty"));
            state.commands.push_back(command(0, "AppendChild"));
            assert!(state.take_ready_commands().is_empty());
        }
        state.batch_depth = 0;
        let commands = state.take_ready_commands();
        assert_eq!(commands.len(), 42 * 4);
        for cell in commands.chunks_exact(4) {
            assert_eq!(cell[0].operation, "CreateElement");
            assert_eq!(cell[3].operation, "AppendChild");
        }
        assert!(state.take_ready_commands().is_empty());
    }

    #[test]
    fn read_flushes_earlier_mutations_in_order() {
        let mut state = State::default();
        state.batch_depth = 1;
        state.commands.push_back(command(0, "CreateElement"));
        state.commands.push_back(command(0, "AppendChild"));
        state.commands.push_back(command(7, "QuerySelector"));
        state.flush_requested = true;
        let operations: Vec<_> = state
            .take_ready_commands()
            .into_iter()
            .map(|command| command.operation)
            .collect();
        assert_eq!(
            operations,
            ["CreateElement", "AppendChild", "QuerySelector"]
        );
        state.commands.push_back(command(0, "SetStyleProperty"));
        assert!(state.take_ready_commands().is_empty());
    }

    #[test]
    fn browser_event_wakes_idle_vm() {
        let session = BrowserSession {
            address: "127.0.0.1:0".parse().unwrap(),
            token: String::new(),
            state: Arc::new(Mutex::new(State::default())),
            event_ready: Arc::new(Condvar::new()),
            wake: Arc::new(tokio::sync::Notify::new()),
        };
        let state = Arc::clone(&session.state);
        let ready = Arc::clone(&session.event_ready);
        let producer = std::thread::spawn(move || {
            std::thread::sleep(Duration::from_millis(10));
            state.lock().unwrap().events.push_back(BrowserEvent {
                document: 1,
                node: 2,
                path: vec![2],
                kind: "click".into(),
                fields: serde_json::Value::Null,
            });
            ready.notify_one();
        });
        let start = Instant::now();
        session.wait_for_event(Duration::from_secs(1));
        assert!(start.elapsed() < Duration::from_millis(500));
        assert_eq!(session.drain_events()[0].kind, "click");
        producer.join().unwrap();
    }
}
