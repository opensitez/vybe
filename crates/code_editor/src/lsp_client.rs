use std::process::{Command, Stdio, Child};
use std::thread;
use std::sync::Arc;
use crossbeam_channel::{Sender, Receiver};
use lsp_server::{Connection, Message, Request, RequestId, Response, Notification};
use lsp_types::{
    InitializeParams, InitializedParams, ClientCapabilities, TraceValue, 
    TextDocumentItem, DidOpenTextDocumentParams, TextDocumentContentChangeEvent, 
    DidChangeTextDocumentParams, VersionedTextDocumentIdentifier, Url,
    Diagnostic, PublishDiagnosticsParams, Position, Range, DiagnosticSeverity
};
use serde_json::to_value;

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
            let dummy_url = Url::parse("file:///Users/youness/www/html/vybe/test.rs").unwrap();

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
                                    let (conn, _) = Connection::new(stdin, stdout);
                                    let params = InitializeParams {
                                        process_id: Some(std::process::id()),
                                        client_info: None,
                                        root_uri: Some(Url::parse("file:///").unwrap()),
                                        initialization_options: None,
                                        capabilities: ClientCapabilities::default(),
                                        trace: Some(TraceValue::Messages),
                                        workspace_folders: None,
                                        locale: None,
                                        root_path: None,
                                    };
                                    let id = conn.send_request::<lsp_types::request::Initialize>(params);
                                    if let Ok(Message::Response(res)) = conn.receiver.recv() {
                                        if res.id == id {
                                            conn.send_notification::<lsp_types::notification::Initialized>(InitializedParams {});
                                            version = 1;
                                            conn.send_notification::<lsp_types::notification::DidOpenTextDocument>(DidOpenTextDocumentParams {
                                                text_document: TextDocumentItem {
                                                    uri: dummy_url.clone(),
                                                    language_id: "rust".to_string(),
                                                    version: version,
                                                    text: content.clone(),
                                                }
                                            });
                                            external_conn = Some(conn);
                                        }
                                    }
                                }
                            } else {
                                // Internal Initialization (e.g. for VB, JS, C#)
                                // We just run the analysis immediately
                                run_internal_analysis(&current_lang, &content, &evt_tx);
                            }
                        }
                        LspRequest::Change(content) => {
                            if let Some(conn) = &external_conn {
                                version += 1;
                                conn.send_notification::<lsp_types::notification::DidChangeTextDocument>(DidChangeTextDocumentParams {
                                    text_document: VersionedTextDocumentIdentifier { uri: dummy_url.clone(), version },
                                    content_changes: vec![TextDocumentContentChangeEvent { range: None, range_length: None, text: content }],
                                });
                            } else {
                                run_internal_analysis(&current_lang, &content, &evt_tx);
                            }
                        }
                        LspRequest::SetLanguage(lang) => {
                            current_lang = lang;
                            external_conn = None; // Kill rust-analyzer if switching
                        }
                        _ => {}
                    }
                }

                // 2. Poll External LSP for Notifications
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
                        let (start, _) = pe.line_col();
                        diagnostics.push(Diagnostic {
                            range: Range::new(Position::new(start.0 as u32 - 1, start.1 as u32 - 1), Position::new(start.0 as u32 - 1, (start.1 as u32 + 5))),
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
                // Parse error message like "Parse error at line 5: ..."
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
