use std::process::{Command, Stdio};
use std::thread;
use std::sync::Arc;
use crossbeam_channel::{Sender, Receiver};
use lsp_server::{Connection, Message, Notification};
use lsp_types::{
    InitializeParams, InitializedParams, ClientCapabilities, TraceValue, 
    TextDocumentItem, DidOpenTextDocumentParams, TextDocumentContentChangeEvent, 
    DidChangeTextDocumentParams, VersionedTextDocumentIdentifier,
    Diagnostic, PublishDiagnosticsParams, Position, Range, DiagnosticSeverity
};
// no direct Url/Uri import; use JSON strings for URIs to avoid type mismatches

pub enum LspEvent {
    Diagnostics(Vec<Diagnostic>),
    Hover(String),
    Definition(Position),
}

pub enum LspRequest {
    Init(String, String), // content, language_id
    Change(String),        // content
    SetLanguage(String),   // language_id
    Hover(u32, u32),       // line, col
    Definition(u32, u32),  // line, col
}

pub struct LspClient {
    tx: Sender<LspRequest>,
    pub rx: Receiver<LspEvent>,
}

impl LspClient {
    pub fn new() -> Self {
        let (req_tx, req_rx) = crossbeam_channel::unbounded();
        let (evt_tx, evt_rx) = crossbeam_channel::unbounded();

        thread::spawn(move || {
            let mut current_lang = "rust".to_string();
            let mut external_conn: Option<Connection> = None;
            let mut version = 0;
            let dummy_uri = "file:///Users/youness/www/html/vybe/test.rs".to_string();

            loop {
                // 1. Handle Requests from Editor
                while let Ok(req) = req_rx.try_recv() {
                    match req {
                        LspRequest::Init(content, lang) => {
                            current_lang = lang;
                            if current_lang == "rust" {
                                let mut cmd = Command::new("rustup");
                                cmd.args(&["run", "stable", "rust-analyzer"]);
                                cmd.stdin(Stdio::piped()).stdout(Stdio::piped()).stderr(Stdio::inherit());
                                if let Ok(mut child) = cmd.spawn() {
                                    let stdin = child.stdin.take().unwrap();
                                    let stdout = child.stdout.take().unwrap();
                                    let conn = Connection::stdio(); // Correct standard method for lsp-server
                                    // Actually, Connection::stdio() takes stdin/stdout, but the crate version we have might be different.
                                    // Based on lsp-server common patterns:
                                    let (conn, _io_threads) = Connection::stdio();
                                    
                                    let params_val = serde_json::json!({
                                        "processId": std::process::id(),
                                        "clientInfo": serde_json::Value::Null,
                                        "rootUri": "file:///",
                                        "initializationOptions": serde_json::Value::Null,
                                        "capabilities": serde_json::to_value(ClientCapabilities::default()).unwrap(),
                                        "trace": "messages",
                                        "workspaceFolders": serde_json::Value::Null,
                                        "locale": serde_json::Value::Null,
                                        "rootPath": serde_json::Value::Null,
                                        "workDoneProgressParams": serde_json::Value::Object(serde_json::Map::new()),
                                    });
                                    let id = conn.sender.send(Message::Request(lsp_server::Request {
                                        id: 1.into(),
                                        method: "initialize".to_string(),
                                        params: params_val,
                                    })).ok();
                                    
                                    if let Ok(Message::Response(res)) = conn.receiver.recv() {
                                        if res.id == 1.into() {
                                            conn.sender.send(Message::Notification(Notification {
                                                method: "initialized".to_string(),
                                                params: serde_json::to_value(InitializedParams {}).unwrap(),
                                            })).ok();
                                            version = 1;
                                            conn.sender.send(Message::Notification(Notification {
                                                method: "textDocument/didOpen".to_string(),
                                                params: serde_json::json!({
                                                    "textDocument": {
                                                        "uri": dummy_uri.clone(),
                                                        "languageId": "rust",
                                                        "version": version,
                                                        "text": content.clone()
                                                    }
                                                }),
                                            })).ok();
                                            external_conn = Some(conn);
                                        }
                                    }
                                }
                            } else {
                                run_internal_analysis(&current_lang, &content, &evt_tx);
                            }
                        }
                        LspRequest::Change(content) => {
                            if let Some(conn) = &external_conn {
                                version += 1;
                                conn.sender.send(Message::Notification(Notification {
                                    method: "textDocument/didChange".to_string(),
                                    params: serde_json::json!({
                                        "textDocument": { "uri": dummy_uri.clone(), "version": version },
                                        "contentChanges": [{ "text": content }]
                                    }),
                                })).ok();
                            } else {
                                run_internal_analysis(&current_lang, &content, &evt_tx);
                            }
                        }
                        LspRequest::SetLanguage(lang) => {
                            current_lang = lang;
                            external_conn = None;
                        }
                        _ => {}
                    }
                }

                if let Some(conn) = &external_conn {
                    while let Ok(msg) = conn.receiver.try_recv() {
                        match msg {
                            Message::Notification(not) if not.method == "textDocument/publishDiagnostics" => {
                                if let Ok(params) = serde_json::from_value::<PublishDiagnosticsParams>(not.params) {
                                    evt_tx.send(LspEvent::Diagnostics(params.diagnostics)).ok();
                                }
                            }
                            _ => {}
                        }
                    }
                }

                thread::sleep(std::time::Duration::from_millis(100));
            }
        });

        Self { tx: req_tx, rx: evt_rx }
    }

    pub fn send(&self, req: LspRequest) { self.tx.send(req).ok(); }
}

fn run_internal_analysis(lang: &str, content: &str, tx: &Sender<LspEvent>) {
    let mut diagnostics = Vec::new();
    match lang {
        "vb" | "basic" => {
            if let Err(e) = vybe_parser_basic::parse_program(content) {
                    match e {
                    vybe_parser_basic::ParseError::PestError(pe) => {
                        // pest::error::LineColLocation can be Pos((line,col)) or Span((sline,scol),(eline,ecol))
                        let (start_line, start_col) = match pe.line_col {
                            pest::error::LineColLocation::Pos((l, c)) => (l, c),
                            pest::error::LineColLocation::Span((l, c), _end) => (l, c),
                        };
                        let sl = start_line.saturating_sub(1) as u32;
                        let sc = start_col.saturating_sub(1) as u32;
                        diagnostics.push(Diagnostic {
                            range: Range::new(Position::new(sl, sc), Position::new(sl, sc + 5)),
                            severity: Some(DiagnosticSeverity::ERROR),
                            code: None,
                            code_description: None,
                            source: Some("vybe-basic".to_string()),
                            message: pe.to_string(),
                            related_information: None,
                            tags: None,
                            data: None,
                        });
                    }
                    _ => {}
                }
            }
        }
        "javascript" | "js" => {
            if let Err(msg) = vybe_parser_js::parse(content) {
                let line = msg.split("line ").nth(1).and_then(|s| s.split(':').next()).and_then(|s| s.parse::<u32>().ok()).unwrap_or(1).saturating_sub(1);
                diagnostics.push(Diagnostic {
                    range: Range::new(Position::new(line, 0), Position::new(line, 80)),
                    severity: Some(DiagnosticSeverity::ERROR),
                    code: None,
                    code_description: None,
                    source: Some("vybe-js".to_string()),
                    message: msg,
                    related_information: None,
                    tags: None,
                    data: None,
                });
            }
        }
        "csharp" | "cs" => {
            if let Err(msg) = vybe_parser_csharp::parse(content) {
                diagnostics.push(Diagnostic {
                    range: Range::new(Position::new(0, 0), Position::new(0, 80)),
                    severity: Some(DiagnosticSeverity::ERROR),
                    code: None,
                    code_description: None,
                    source: Some("vybe-cs".to_string()),
                    message: msg,
                    related_information: None,
                    tags: None,
                    data: None,
                });
            }
        }
        _ => {}
    }
    tx.send(LspEvent::Diagnostics(diagnostics)).ok();
}
