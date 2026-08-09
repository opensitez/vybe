use crossbeam_channel::{Receiver, Sender};
use lsp_types::{CompletionItemKind, Diagnostic, DiagnosticSeverity, Position, Range};
use std::collections::HashMap;
use std::thread;
use vybe_compiler::lsp::{AnalysisResult, SymbolKind};

#[derive(Debug, Clone)]
pub struct SimpleCompletion {
    pub label: String,
    pub detail: Option<String>,
    pub insert_text: String,
    pub kind: Option<CompletionItemKind>,
}

pub enum LspEvent {
    Diagnostics(String, Vec<Diagnostic>),
    Completion(Vec<SimpleCompletion>),
    #[allow(dead_code)]
    Hover(String, String),
    #[allow(dead_code)]
    Definition(String, Position),
}

#[allow(dead_code)]
pub enum LspRequest {
    Init(String, String, String), // content, language_id, uri
    Change(String, String),       // content, uri
    Completion(String, u32, u32), // uri, line, col
    #[allow(dead_code)]
    Close(String),
    #[allow(dead_code)]
    Hover(String, u32, u32),
    #[allow(dead_code)]
    Definition(String, u32, u32),
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
            // uri → (content, analysis result)
            let mut cache: HashMap<String, (String, AnalysisResult)> = HashMap::new();

            loop {
                crossbeam_channel::select! {
                    recv(req_rx) -> req => {
                        let Ok(req) = req else { continue };
                        match req {
                            LspRequest::Init(content, _, uri) | LspRequest::Change(content, uri) => {
                                let result = vybe_compiler::lsp::analyze(&uri, &content);
                                emit_diagnostics(&result, &uri, &evt_tx);
                                cache.insert(uri, (content, result));
                            }
                            LspRequest::Completion(uri, line, col) => {
                                if let Some((content, result)) = cache.get(&uri) {
                                    let prefix = word_before(content, line, col);
                                    evt_tx.send(LspEvent::Completion(build_completions(result, &prefix))).ok();
                                }
                            }
                            LspRequest::Hover(uri, line, col) => {
                                if let Some((content, result)) = cache.get(&uri) {
                                    let word = word_before(content, line, col);
                                    if let Some(sym) = find_symbol(&result.symbols, &word) {
                                        let text = if sym.detail.is_empty() {
                                            sym.name.clone()
                                        } else {
                                            format!("{} {}", sym.name, sym.detail)
                                        };
                                        evt_tx.send(LspEvent::Hover(uri, text)).ok();
                                    }
                                }
                            }
                            LspRequest::Definition(uri, line, col) => {
                                if let Some((content, result)) = cache.get(&uri) {
                                    let word = word_before(content, line, col);
                                    if let Some(sym) = find_symbol(&result.symbols, &word) {
                                        evt_tx.send(LspEvent::Definition(
                                            uri.clone(),
                                            Position::new(sym.line, 0),
                                        )).ok();
                                    }
                                }
                            }
                            LspRequest::Close(_) => {}
                        }
                    }
                }
            }
        });

        Self {
            tx: req_tx,
            rx: evt_rx,
        }
    }

    pub fn send(&self, req: LspRequest) {
        self.tx.send(req).ok();
    }
}

// ── Helpers ────────────────────────────────────────────────────────────────

fn emit_diagnostics(result: &AnalysisResult, uri: &str, tx: &Sender<LspEvent>) {
    let diagnostics = result
        .diagnostics
        .iter()
        .map(|d| Diagnostic {
            range: Range::new(
                Position::new(d.line, d.col),
                Position::new(d.line, d.end_col),
            ),
            severity: Some(match d.severity {
                vybe_compiler::lsp::DiagSeverity::Error => DiagnosticSeverity::ERROR,
                vybe_compiler::lsp::DiagSeverity::Warning => DiagnosticSeverity::WARNING,
                vybe_compiler::lsp::DiagSeverity::Info => DiagnosticSeverity::INFORMATION,
            }),
            source: Some("vybe".to_string()),
            message: d.message.clone(),
            ..Default::default()
        })
        .collect();
    tx.send(LspEvent::Diagnostics(uri.to_string(), diagnostics))
        .ok();
}

/// Word immediately before (line, col) — the completion prefix.
fn word_before(content: &str, line: u32, col: u32) -> String {
    let line_text = content.lines().nth(line as usize).unwrap_or("");
    let col = (col as usize).min(line_text.len());
    let start = line_text[..col]
        .rfind(|c: char| !c.is_alphanumeric() && c != '_')
        .map(|i| i + 1)
        .unwrap_or(0);
    line_text[start..col].to_string()
}

/// Recursively find a symbol by name (case-insensitive).
fn find_symbol<'a>(
    syms: &'a [vybe_compiler::lsp::Symbol],
    name: &str,
) -> Option<&'a vybe_compiler::lsp::Symbol> {
    if name.is_empty() {
        return None;
    }
    let lower = name.to_lowercase();
    for sym in syms {
        if sym.name.to_lowercase() == lower {
            return Some(sym);
        }
        if let Some(f) = find_symbol(&sym.children, name) {
            return Some(f);
        }
    }
    None
}

fn build_completions(result: &AnalysisResult, prefix: &str) -> Vec<SimpleCompletion> {
    let lower = prefix.to_lowercase();
    let mut items = Vec::new();
    collect_from_symbols(&result.symbols, &lower, &mut items);
    for &kw in result.keywords {
        if kw.to_lowercase().starts_with(&lower) {
            items.push(SimpleCompletion {
                label: kw.to_string(),
                detail: None,
                insert_text: kw.to_string(),
                kind: Some(CompletionItemKind::KEYWORD),
            });
        }
    }
    items.sort_by(|a, b| a.label.cmp(&b.label));
    items.dedup_by(|a, b| a.label == b.label);
    items.truncate(60);
    items
}

fn collect_from_symbols(
    syms: &[vybe_compiler::lsp::Symbol],
    prefix: &str,
    out: &mut Vec<SimpleCompletion>,
) {
    for sym in syms {
        if sym.name.to_lowercase().starts_with(prefix) {
            out.push(SimpleCompletion {
                label: sym.name.clone(),
                detail: if sym.detail.is_empty() {
                    None
                } else {
                    Some(sym.detail.clone())
                },
                insert_text: sym.name.clone(),
                kind: Some(kind_to_lsp(sym.kind)),
            });
        }
        collect_from_symbols(&sym.children, prefix, out);
    }
}

fn kind_to_lsp(k: SymbolKind) -> CompletionItemKind {
    match k {
        SymbolKind::Function | SymbolKind::Procedure => CompletionItemKind::FUNCTION,
        SymbolKind::Class => CompletionItemKind::CLASS,
        SymbolKind::Interface => CompletionItemKind::INTERFACE,
        SymbolKind::Variable => CompletionItemKind::VARIABLE,
        SymbolKind::Constant => CompletionItemKind::CONSTANT,
        SymbolKind::Field => CompletionItemKind::FIELD,
        SymbolKind::Property => CompletionItemKind::PROPERTY,
        SymbolKind::Method | SymbolKind::Constructor => CompletionItemKind::METHOD,
        SymbolKind::Module => CompletionItemKind::MODULE,
        SymbolKind::Enum => CompletionItemKind::ENUM,
        SymbolKind::EnumMember => CompletionItemKind::ENUM_MEMBER,
        SymbolKind::Struct => CompletionItemKind::STRUCT,
        SymbolKind::Event => CompletionItemKind::EVENT,
        SymbolKind::Type => CompletionItemKind::TYPE_PARAMETER,
    }
}
