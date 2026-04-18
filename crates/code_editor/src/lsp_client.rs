use std::process::{Command, Stdio};
use std::sync::atomic::{AtomicI32, Ordering};
use std::sync::{Arc, Mutex};
use std::thread;
use crossbeam_channel::{Sender, Receiver};
use lsp_server::{Message, Notification};
use lsp_types::{
    InitializedParams, ClientCapabilities, 
    Diagnostic, PublishDiagnosticsParams, Position, Range, DiagnosticSeverity,
    CompletionResponse, CompletionItemKind, Hover, GotoDefinitionResponse,
};
// no direct Url/Uri import; use JSON strings for URIs to avoid type mismatches

/// A simplified completion item for the UI layer.
#[derive(Debug, Clone)]
pub struct SimpleCompletion {
    pub label: String,
    pub detail: Option<String>,
    pub insert_text: String,
    pub kind: Option<CompletionItemKind>,
}

pub enum LspEvent {
    Diagnostics(String, Vec<Diagnostic>), // URI, Diagnostics
    Completion(Vec<SimpleCompletion>),     // Completion items
    #[allow(dead_code)]
    Hover(String, String),               // URI, Hover text
    #[allow(dead_code)]
    Definition(String, Position),        // URI, Position
}

pub enum LspRequest {
    Init(String, String, String), // content, language_id, uri
    Change(String, String),        // content, uri
    Completion(String, u32, u32),  // uri, line, col
    #[allow(dead_code)]
    Close(String),                 // uri
    #[allow(dead_code)]
    Hover(String, u32, u32),       // uri, line, col
    #[allow(dead_code)]
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
            let next_id = Arc::new(AtomicI32::new(2)); // 1 is reserved for init
            let pending: Arc<Mutex<std::collections::HashMap<i32, String>>> = Arc::new(Mutex::new(std::collections::HashMap::new()));

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
                                                let pending_clone = pending.clone();
                                                thread::spawn(move || {
                                                    while let Ok(Some(msg)) = Message::read(&mut stdout) {
                                                        match msg {
                                                            Message::Notification(not) if not.method == "textDocument/publishDiagnostics" => {
                                                                if let Ok(params) = serde_json::from_value::<PublishDiagnosticsParams>(not.params) {
                                                                    etx.send(LspEvent::Diagnostics(params.uri.to_string(), params.diagnostics)).ok();
                                                                }
                                                            }
                                                            Message::Response(res) if res.id == 1.into() => {
                                                                // Init response — ignored
                                                            }
                                                            Message::Response(res) => {
                                                                let id_num = res.id.to_string().parse::<i32>().unwrap_or(0);
                                                                let method = pending_clone.lock().ok()
                                                                    .and_then(|mut m| m.remove(&id_num));
                                                                if let Some(result) = res.result {
                                                                    match method.as_deref() {
                                                                        Some("textDocument/completion") => {
                                                                            if let Ok(cr) = serde_json::from_value::<CompletionResponse>(result) {
                                                                                let items = match cr {
                                                                                    CompletionResponse::Array(arr) => arr,
                                                                                    CompletionResponse::List(list) => list.items,
                                                                                };
                                                                                let simple: Vec<SimpleCompletion> = items.into_iter().map(|ci| {
                                                                                    SimpleCompletion {
                                                                                        label: ci.label.clone(),
                                                                                        detail: ci.detail.clone(),
                                                                                        insert_text: ci.insert_text.unwrap_or(ci.label),
                                                                                        kind: ci.kind,
                                                                                    }
                                                                                }).collect();
                                                                                etx.send(LspEvent::Completion(simple)).ok();
                                                                            }
                                                                        }
                                                                        Some("textDocument/hover") => {
                                                                            if let Ok(hover) = serde_json::from_value::<Hover>(result) {
                                                                                let text = match hover.contents {
                                                                                    lsp_types::HoverContents::Scalar(mc) => match mc {
                                                                                        lsp_types::MarkedString::String(s) => s,
                                                                                        lsp_types::MarkedString::LanguageString(ls) => ls.value,
                                                                                    },
                                                                                    lsp_types::HoverContents::Array(arr) => arr.into_iter().map(|mc| match mc {
                                                                                        lsp_types::MarkedString::String(s) => s,
                                                                                        lsp_types::MarkedString::LanguageString(ls) => ls.value,
                                                                                    }).collect::<Vec<_>>().join("\n"),
                                                                                    lsp_types::HoverContents::Markup(mc) => mc.value,
                                                                                };
                                                                                if !text.is_empty() {
                                                                                    etx.send(LspEvent::Hover(String::new(), text)).ok();
                                                                                }
                                                                            }
                                                                        }
                                                                        Some("textDocument/definition") => {
                                                                            if let Ok(def) = serde_json::from_value::<GotoDefinitionResponse>(result) {
                                                                                let (uri, pos) = match def {
                                                                                    GotoDefinitionResponse::Scalar(loc) => (loc.uri.to_string(), loc.range.start),
                                                                                    GotoDefinitionResponse::Array(locs) => {
                                                                                        if let Some(loc) = locs.first() { (loc.uri.to_string(), loc.range.start) }
                                                                                        else { return; }
                                                                                    }
                                                                                    GotoDefinitionResponse::Link(links) => {
                                                                                        if let Some(link) = links.first() { (link.target_uri.to_string(), link.target_selection_range.start) }
                                                                                        else { return; }
                                                                                    }
                                                                                };
                                                                                etx.send(LspEvent::Definition(uri, pos)).ok();
                                                                            }
                                                                        }
                                                                        _ => {}
                                                                    }
                                                                }
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
                                LspRequest::Completion(uri, line, col) => {
                                    send_lsp_request(&mut child_in, &next_id, &pending, "textDocument/completion", &uri, line, col);
                                }
                                LspRequest::Hover(uri, line, col) => {
                                    send_lsp_request(&mut child_in, &next_id, &pending, "textDocument/hover", &uri, line, col);
                                }
                                LspRequest::Definition(uri, line, col) => {
                                    send_lsp_request(&mut child_in, &next_id, &pending, "textDocument/definition", &uri, line, col);
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
    let result = vybe_lsp::analyze(uri, content);
    let diagnostics = result.diagnostics.into_iter().map(|d| {
        Diagnostic {
            range: Range::new(
                Position::new(d.line, d.col),
                Position::new(d.line, d.end_col),
            ),
            severity: Some(match d.severity {
                vybe_lsp::DiagSeverity::Error => DiagnosticSeverity::ERROR,
                vybe_lsp::DiagSeverity::Warning => DiagnosticSeverity::WARNING,
                vybe_lsp::DiagSeverity::Info => DiagnosticSeverity::INFORMATION,
            }),
            code: None,
            code_description: None,
            source: Some(format!("vybe-{}", lang)),
            message: d.message,
            related_information: None,
            tags: None,
            data: None,
        }
    }).collect();
    tx.send(LspEvent::Diagnostics(uri.to_string(), diagnostics)).ok();
}

fn send_lsp_request(
    child_in: &mut Option<std::io::BufWriter<std::process::ChildStdin>>,
    next_id: &Arc<AtomicI32>,
    pending: &Arc<Mutex<std::collections::HashMap<i32, String>>>,
    method: &str,
    uri: &str,
    line: u32,
    col: u32,
) {
    if let Some(stdin) = child_in.as_mut() {
        let id = next_id.fetch_add(1, Ordering::Relaxed);
        if let Ok(mut map) = pending.lock() { map.insert(id, method.to_string()); }
        let req = lsp_server::Request {
            id: id.into(),
            method: method.to_string(),
            params: serde_json::json!({
                "textDocument": { "uri": uri },
                "position": { "line": line, "character": col },
            }),
        };
        Message::Request(req).write(stdin).ok();
    }
}
