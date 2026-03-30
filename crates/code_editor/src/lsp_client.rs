use std::process::{Command, Stdio};
use std::thread;
use crossbeam_channel::{Sender, Receiver};
use lsp_server::{Message, Notification};
use lsp_types::{
    InitializedParams, ClientCapabilities, 
    Diagnostic, PublishDiagnosticsParams, Position, Range, DiagnosticSeverity
};
// no direct Url/Uri import; use JSON strings for URIs to avoid type mismatches

pub enum LspEvent {
    Diagnostics(String, Vec<Diagnostic>), // URI, Diagnostics
    Hover(String, String),               // URI, Hover text
    Definition(String, Position),        // URI, Position
}

pub enum LspRequest {
    Init(String, String, String), // content, language_id, uri
    Change(String, String),        // content, uri
    Close(String),                 // uri
    Hover(String, u32, u32),       // uri, line, col
    Definition(String, u32, u32),  // uri, line, col
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
            let mut child: Option<std::process::Child> = None;
            let mut child_in: Option<std::io::BufWriter<std::process::ChildStdin>> = None;
            let mut versions: std::collections::HashMap<String, i32> = std::collections::HashMap::new();

            loop {
                crossbeam_channel::select! {
                    recv(req_rx) -> req => {
                        if let Ok(req) = req {
                            match req {
                                LspRequest::Init(content, lang, uri) => {
                                    if lang == "rust" && child.is_none() {
                                        if let Ok(mut c) = Command::new("rust-analyzer")
                                            .stdin(Stdio::piped())
                                            .stdout(Stdio::piped())
                                            .stderr(Stdio::inherit())
                                            .spawn() 
                                        {
                                            let stdin = std::io::BufWriter::new(c.stdin.take().unwrap());
                                            let mut stdout = std::io::BufReader::new(c.stdout.take().unwrap());
                                            child_in = Some(stdin);
                                            child = Some(c);

                                            // Initialization sequence
                                            if let Some(mut stdin) = child_in.as_mut() {
                                                let params_val = serde_json::json!({
                                                    "processId": std::process::id(),
                                                    "capabilities": serde_json::to_value(ClientCapabilities::default()).unwrap(),
                                                    "rootUri": format!("file://{}", std::env::current_dir().unwrap_or_default().to_string_lossy()),
                                                });
                                                let init_req = lsp_server::Request { id: 1.into(), method: "initialize".to_string(), params: params_val };
                                                Message::Request(init_req).write(&mut stdin).ok();
                                                
                                                // Reader thread for this child
                                                let etx = evt_tx.clone();
                                                thread::spawn(move || {
                                                    while let Ok(Some(msg)) = Message::read(&mut stdout) {
                                                        match msg {
                                                            Message::Notification(not) if not.method == "textDocument/publishDiagnostics" => {
                                                                if let Ok(params) = serde_json::from_value::<PublishDiagnosticsParams>(not.params) {
                                                                    etx.send(LspEvent::Diagnostics(params.uri.to_string(), params.diagnostics)).ok();
                                                                }
                                                            }
                                                            Message::Response(res) if res.id == 1.into() => {
                                                                // Initialized will be sent next
                                                            }
                                                            _ => {}
                                                        }
                                                    }
                                                });

                                                // Send initialized
                                                Message::Notification(Notification { 
                                                    method: "initialized".to_string(), 
                                                    params: serde_json::to_value(InitializedParams {}).unwrap() 
                                                }).write(&mut stdin).ok();
                                            }
                                        }
                                    }
                                    
                                    if let Some(mut stdin) = child_in.as_mut() {
                                        let v = versions.entry(uri.clone()).or_insert(0);
                                        *v += 1;
                                        Message::Notification(Notification { 
                                            method: "textDocument/didOpen".to_string(), 
                                            params: serde_json::json!({ 
                                                "textDocument": { "uri": uri, "languageId": lang, "version": *v, "text": content } 
                                            }) 
                                        }).write(&mut stdin).ok();
                                    } else {
                                        run_internal_analysis(&lang, &content, &uri, &evt_tx);
                                    }
                                }
                                LspRequest::Change(content, uri) => {
                                    if let Some(mut stdin) = child_in.as_mut() {
                                        let v = versions.entry(uri.clone()).or_insert(0);
                                        *v += 1;
                                        Message::Notification(Notification { 
                                            method: "textDocument/didChange".to_string(), 
                                            params: serde_json::json!({ 
                                                "textDocument": { "uri": uri, "version": *v }, 
                                                "contentChanges": [{ "text": content }] 
                                            }) 
                                        }).write(&mut stdin).ok();
                                    }
                                }
                                _ => {}
                            }
                        }
                    }
                }
            }
        });

        Self { tx: req_tx, rx: evt_rx }
    }

    pub fn send(&self, req: LspRequest) { self.tx.send(req).ok(); }
}

fn run_internal_analysis(lang: &str, content: &str, uri: &str, tx: &Sender<LspEvent>) {
    let mut diagnostics = Vec::new();
    match lang {
        "vb" | "basic" => {
            if let Err(e) = vybe_parser_basic::parse_program(content) {
                    match e {
                    vybe_parser_basic::ParseError::PestError(pe) => {
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
    tx.send(LspEvent::Diagnostics(uri.to_string(), diagnostics)).ok();
}
