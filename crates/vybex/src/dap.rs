//! Debug Adapter Protocol (DAP) server — lets VS Code (or any DAP client)
//! attach to a running Vybe program with its native breakpoint UI, step
//! buttons, call-stack, variables pane, and debug console.
//!
//! It is a thin TRANSLATION layer: it speaks DAP JSON over a TCP socket and
//! maps each request onto the same typed `DebugCommand`/`DebugEvent` protocol
//! the built-in REPL uses (`debug_repl.rs`). The VM stays on the main thread;
//! this runs on worker threads holding only channel endpoints.
//!
//! Launch: `vybex <file> --dap-port 4711`, then a VS Code launch config with
//! `"debugServer": 4711`. The program pauses on entry until the client sends
//! `configurationDone` / `continue`.

use std::io::{BufReader, Read, Write};
use std::net::TcpListener;
use std::sync::atomic::{AtomicI64, Ordering};
use std::sync::mpsc::{channel, Receiver, Sender};
use std::sync::{Arc, Mutex};
use std::thread;

use serde_json::{json, Value as J};

use vybe_bytecode::debugger::{DebugEvent, DebugResponse, PauseReason};
use vybe_bytecode::{DebugCommand, DebugRequest, VM};

/// Attach a DAP server to `vm` (pausing on entry) and spawn the TCP listener.
/// `source_path` is reported to the client for stack frames.
pub fn attach(vm: &mut VM, port: u16, source_path: String) {
    let (cmd_tx, cmd_rx) = channel::<DebugRequest>();
    let (evt_tx, evt_rx) = channel::<DebugEvent>();
    vm.attach_debugger(cmd_rx, evt_tx, /* pause_on_entry */ true);
    thread::spawn(move || {
        if let Err(e) = serve(port, cmd_tx, evt_rx, source_path) {
            eprintln!("[dap] server error: {e}");
        }
    });
}

fn serve(
    port: u16,
    cmd_tx: Sender<DebugRequest>,
    evt_rx: Receiver<DebugEvent>,
    source_path: String,
) -> std::io::Result<()> {
    let listener = TcpListener::bind(("127.0.0.1", port))?;
    eprintln!("[dap] listening on 127.0.0.1:{port} — connect VS Code with \"debugServer\": {port}");
    let (stream, peer) = listener.accept()?;
    eprintln!("[dap] client connected: {peer}");

    let seq = Arc::new(AtomicI64::new(1));
    let writer = Arc::new(Mutex::new(stream.try_clone()?));

    // Event pump: DebugEvent → DAP events.
    {
        let writer = writer.clone();
        let seq = seq.clone();
        thread::spawn(move || {
            for event in evt_rx {
                emit_event(&writer, &seq, event);
            }
        });
    }

    // Request loop: DAP requests → DebugCommand → DAP responses.
    let session = Session { cmd_tx, writer, seq, source_path };
    let mut reader = BufReader::new(stream);
    while let Some(msg) = read_message(&mut reader) {
        if msg.get("type").and_then(J::as_str) == Some("request") {
            if session.handle_request(&msg) {
                break; // disconnect / terminate
            }
        }
    }
    eprintln!("[dap] client disconnected");
    Ok(())
}

struct Session {
    cmd_tx: Sender<DebugRequest>,
    writer: Arc<Mutex<std::net::TcpStream>>,
    seq: Arc<AtomicI64>,
    source_path: String,
}

impl Session {
    /// Handle one DAP request. Returns true if the session should end.
    fn handle_request(&self, req: &J) -> bool {
        let command = req.get("command").and_then(J::as_str).unwrap_or("");
        let req_seq = req.get("seq").and_then(J::as_i64).unwrap_or(0);
        let args = req.get("arguments").cloned().unwrap_or(J::Null);

        match command {
            "initialize" => {
                self.respond(req_seq, command, true, json!({
                    "supportsConfigurationDoneRequest": true,
                    "supportsConditionalBreakpoints": true,
                    "supportsHitConditionalBreakpoints": true,
                    "supportsLogPoints": true,
                    "supportsSetVariable": true,
                    "supportsEvaluateForHovers": true,
                    "supportsTerminateRequest": true,
                    "supportsRestartRequest": true,
                }));
                // Ready for breakpoint configuration.
                self.event("initialized", J::Null);
            }
            "launch" | "attach" => {
                self.respond(req_seq, command, true, J::Null);
            }
            "setBreakpoints" => {
                // Replace this source's breakpoints (single-file model: clear all).
                self.run(DebugCommand::ClearBreakpoints);
                let mut verified = Vec::new();
                if let Some(bps) = args.get("breakpoints").and_then(J::as_array) {
                    for bp in bps {
                        let line = bp.get("line").and_then(J::as_i64).unwrap_or(0) as u32;
                        let condition = bp
                            .get("condition")
                            .and_then(J::as_str)
                            .filter(|s| !s.trim().is_empty())
                            .map(|s| s.to_string());
                        let resp = self.run(DebugCommand::BreakSourceLine { line, condition });
                        let ok = matches!(resp, Some(DebugResponse::Breakpoints(ref v)) if !v.is_empty());
                        let actual_line = match &resp {
                            Some(DebugResponse::Breakpoints(v)) => {
                                v.first().and_then(|b| b.line).unwrap_or(line)
                            }
                            _ => line,
                        };
                        verified.push(json!({ "verified": ok, "line": actual_line }));
                    }
                }
                self.respond(req_seq, command, true, json!({ "breakpoints": verified }));
            }
            "configurationDone" => {
                self.respond(req_seq, command, true, J::Null);
            }
            "threads" => {
                self.respond(req_seq, command, true, json!({
                    "threads": [{ "id": 1, "name": "main" }]
                }));
            }
            "stackTrace" => {
                let frames = match self.run(DebugCommand::Backtrace) {
                    Some(DebugResponse::Backtrace(fs)) => fs
                        .iter()
                        .map(|f| json!({
                            "id": f.depth,
                            "name": f.chunk_name,
                            "line": f.line.unwrap_or(0),
                            "column": 1,
                            "source": { "name": basename(&self.source_path), "path": self.source_path },
                        }))
                        .collect::<Vec<_>>(),
                    _ => Vec::new(),
                };
                let total = frames.len();
                self.respond(req_seq, command, true, json!({
                    "stackFrames": frames, "totalFrames": total
                }));
            }
            "scopes" => {
                let frame_id = args.get("frameId").and_then(J::as_i64).unwrap_or(0);
                // variablesReference encodes the frame (offset by 1 so it's nonzero).
                self.respond(req_seq, command, true, json!({
                    "scopes": [{
                        "name": "Locals",
                        "variablesReference": frame_id + 1,
                        "expensive": false,
                    }]
                }));
            }
            "variables" => {
                let vref = args.get("variablesReference").and_then(J::as_i64).unwrap_or(0);
                let frame = (vref - 1).max(0) as usize;
                let vars = match self.run(DebugCommand::Locals { frame }) {
                    Some(DebugResponse::Locals(slots)) => slots
                        .iter()
                        .map(|s| json!({
                            "name": s.name.clone().unwrap_or_else(|| format!("[{}]", s.index)),
                            "value": s.value,
                            "variablesReference": 0,
                        }))
                        .collect::<Vec<_>>(),
                    _ => Vec::new(),
                };
                self.respond(req_seq, command, true, json!({ "variables": vars }));
            }
            "continue" => {
                self.run(DebugCommand::Continue);
                self.respond(req_seq, command, true, json!({ "allThreadsContinued": true }));
            }
            "next" => {
                self.run(DebugCommand::StepOver);
                self.respond(req_seq, command, true, J::Null);
            }
            "stepIn" => {
                self.run(DebugCommand::StepInto);
                self.respond(req_seq, command, true, J::Null);
            }
            "stepOut" => {
                self.run(DebugCommand::StepOut);
                self.respond(req_seq, command, true, J::Null);
            }
            "pause" => {
                self.run(DebugCommand::Pause);
                self.respond(req_seq, command, true, J::Null);
            }
            "evaluate" => {
                let expr = args.get("expression").and_then(J::as_str).unwrap_or("");
                let (ok, result) = match self.run(DebugCommand::Print { path: expr.to_string() }) {
                    Some(DebugResponse::Value(v)) => (true, v),
                    Some(DebugResponse::Error(e)) => (false, e),
                    _ => (false, "eval unavailable".to_string()),
                };
                self.respond(req_seq, command, ok, json!({
                    "result": result, "variablesReference": 0
                }));
            }
            "setVariable" => {
                let name = args.get("name").and_then(J::as_str).unwrap_or("").to_string();
                let value = args.get("value").and_then(J::as_str).unwrap_or("").to_string();
                let (ok, shown) = match self.run(DebugCommand::SetVar { name, literal: value.clone() }) {
                    Some(DebugResponse::Value(_)) => (true, value),
                    Some(DebugResponse::Error(e)) => (false, e),
                    _ => (false, "set failed".to_string()),
                };
                self.respond(req_seq, command, ok, json!({
                    "value": shown, "variablesReference": 0
                }));
            }
            "disconnect" | "terminate" => {
                self.run(DebugCommand::Quit);
                self.respond(req_seq, command, true, J::Null);
                return true;
            }
            other => {
                // Unknown request — acknowledge so the client isn't blocked.
                self.respond(req_seq, other, true, J::Null);
            }
        }
        false
    }

    /// Send a `DebugCommand` and wait for its reply.
    fn run(&self, command: DebugCommand) -> Option<DebugResponse> {
        let (reply_tx, reply_rx) = channel::<DebugResponse>();
        self.cmd_tx.send(DebugRequest { command, reply: reply_tx }).ok()?;
        reply_rx.recv().ok()
    }

    fn respond(&self, request_seq: i64, command: &str, success: bool, body: J) {
        let msg = json!({
            "seq": self.seq.fetch_add(1, Ordering::SeqCst),
            "type": "response",
            "request_seq": request_seq,
            "success": success,
            "command": command,
            "body": body,
        });
        write_message(&self.writer, &msg);
    }

    fn event(&self, event: &str, body: J) {
        emit_named(&self.writer, &self.seq, event, body);
    }
}

fn emit_event(writer: &Arc<Mutex<std::net::TcpStream>>, seq: &Arc<AtomicI64>, event: DebugEvent) {
    match event {
        DebugEvent::Paused { reason, .. } => {
            emit_named(writer, seq, "stopped", json!({
                "reason": dap_stop_reason(&reason),
                "threadId": 1,
                "allThreadsStopped": true,
            }));
        }
        DebugEvent::Resumed => {
            emit_named(writer, seq, "continued", json!({ "threadId": 1, "allThreadsContinued": true }));
        }
        DebugEvent::Exited { value } => {
            emit_named(writer, seq, "output", json!({ "category": "stdout", "output": format!("exited → {value}\n") }));
            emit_named(writer, seq, "terminated", J::Null);
        }
        DebugEvent::Log { message } => {
            emit_named(writer, seq, "output", json!({ "category": "console", "output": format!("{message}\n") }));
        }
        DebugEvent::Opcode { .. } => {}
    }
}

fn dap_stop_reason(r: &PauseReason) -> &'static str {
    match r {
        PauseReason::Entry => "entry",
        PauseReason::Breakpoint { .. } => "breakpoint",
        PauseReason::Step => "step",
        PauseReason::Interrupt => "pause",
        PauseReason::Watchpoint { .. } => "data breakpoint",
        PauseReason::Exception { .. } => "exception",
    }
}

fn emit_named(writer: &Arc<Mutex<std::net::TcpStream>>, seq: &Arc<AtomicI64>, event: &str, body: J) {
    let msg = json!({
        "seq": seq.fetch_add(1, Ordering::SeqCst),
        "type": "event",
        "event": event,
        "body": body,
    });
    write_message(writer, &msg);
}

// ─── DAP wire framing (`Content-Length: N\r\n\r\n<json>`) ────────────────────

fn write_message(writer: &Arc<Mutex<std::net::TcpStream>>, msg: &J) {
    let body = msg.to_string();
    let framed = format!("Content-Length: {}\r\n\r\n{}", body.len(), body);
    if let Ok(mut w) = writer.lock() {
        let _ = w.write_all(framed.as_bytes());
        let _ = w.flush();
    }
}

fn read_message<R: Read>(reader: &mut BufReader<R>) -> Option<J> {
    use std::io::BufRead;
    // Read headers until a blank line; capture Content-Length.
    let mut content_length = 0usize;
    loop {
        let mut line = String::new();
        if reader.read_line(&mut line).ok()? == 0 {
            return None; // EOF
        }
        let trimmed = line.trim_end();
        if trimmed.is_empty() {
            break;
        }
        if let Some(rest) = trimmed.strip_prefix("Content-Length:") {
            content_length = rest.trim().parse().ok()?;
        }
    }
    if content_length == 0 {
        return None;
    }
    let mut buf = vec![0u8; content_length];
    reader.read_exact(&mut buf).ok()?;
    serde_json::from_slice(&buf).ok()
}

fn basename(path: &str) -> String {
    std::path::Path::new(path)
        .file_name()
        .and_then(|s| s.to_str())
        .unwrap_or(path)
        .to_string()
}
