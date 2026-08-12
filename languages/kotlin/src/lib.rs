// Force-link every plugin crate in `[dependencies]` so its link-time
// registration reaches the registry. Generated from Cargo.toml — see build.rs.
include!(concat!(env!("OUT_DIR"), "/linked_plugins.rs"));

pub mod emitter;
pub mod normalize_class;
pub mod protocol;
pub mod tree_register;
mod walker;

use pest_derive::Parser;

#[derive(Parser)]
#[grammar = "src/grammar.pest"]
pub(crate) struct KotlinParser;

/// Parse Kotlin source into the common AST.
pub fn parse(source: &str) -> Result<vybe_ast::Module, String> {
    // Parse on a thread with a LARGE stack. The expression grammar is ~17
    // levels deep and `walk_expr`'s frames are big in debug builds, so a
    // moderately nested expression (`(xs.zip(listOf(1))).toString()`) walks a
    // pair tree deep enough to blow the default 8 MiB main stack — measured
    // as a hard overflow, not a hang. The documented alternative (flattening
    // the precedence chain) caps how faithful the grammar can be; a worker
    // thread with headroom does not.
    let source = normalize_escaped_identifiers(source);
    std::thread::Builder::new()
        .name("kotlin-parse".into())
        .stack_size(256 * 1024 * 1024)
        .spawn(move || walker::parse(&source))
        .map_err(|e| format!("parse thread: {e}"))?
        .join()
        .map_err(|_| "Kotlin parse thread panicked".to_string())?
}

fn normalize_escaped_identifiers(source: &str) -> String {
    let bytes = source.as_bytes();
    let mut out = String::with_capacity(source.len());
    let mut i = 0;
    while i < bytes.len() {
        match bytes[i] {
            b'`' => {
                let start = i + 1;
                i = start;
                while i < bytes.len() && bytes[i] != b'`' {
                    i += 1;
                }
                let name = &source[start..i.min(bytes.len())];
                out.push_str("__kt_escaped_");
                for byte in name.as_bytes() {
                    use std::fmt::Write;
                    let _ = write!(out, "{byte:02x}");
                }
                if i < bytes.len() {
                    i += 1;
                }
            }
            b'"' if source[i..].starts_with("\"\"\"") => {
                let start = i;
                i += 3;
                while i < bytes.len() && !source[i..].starts_with("\"\"\"") {
                    i += 1;
                }
                if i < bytes.len() {
                    i += 3;
                }
                out.push_str(&source[start..i.min(bytes.len())]);
            }
            b'"' => {
                let start = i;
                i += 1;
                while i < bytes.len() {
                    if bytes[i] == b'\\' {
                        i = (i + 2).min(bytes.len());
                    } else if bytes[i] == b'"' {
                        i += 1;
                        break;
                    } else {
                        i += 1;
                    }
                }
                out.push_str(&source[start..i.min(bytes.len())]);
            }
            b'\'' => {
                let start = i;
                i += 1;
                while i < bytes.len() {
                    if bytes[i] == b'\\' {
                        i = (i + 2).min(bytes.len());
                    } else if bytes[i] == b'\'' {
                        i += 1;
                        break;
                    } else {
                        i += 1;
                    }
                }
                out.push_str(&source[start..i.min(bytes.len())]);
            }
            b'/' if i + 1 < bytes.len() && bytes[i + 1] == b'/' => {
                let start = i;
                i += 2;
                while i < bytes.len() && bytes[i] != b'\n' {
                    i += 1;
                }
                out.push_str(&source[start..i.min(bytes.len())]);
            }
            b'/' if i + 1 < bytes.len() && bytes[i + 1] == b'*' => {
                let start = i;
                i += 2;
                while i + 1 < bytes.len() && !(bytes[i] == b'*' && bytes[i + 1] == b'/') {
                    i += 1;
                }
                if i + 1 < bytes.len() {
                    i += 2;
                }
                out.push_str(&source[start..i.min(bytes.len())]);
            }
            _ => {
                let ch = source[i..].chars().next().unwrap();
                out.push(ch);
                i += ch.len_utf8();
            }
        }
    }
    out
}

/// Embedded profile TOML source.
pub fn profile_source() -> &'static str {
    include_str!("profile")
}

/// Register this language with the shared plugin registry.
pub fn register() {
    vybe_runtime::registry::register_language(vybe_runtime::registry::LanguageDef {
        name: "kotlin",
        parse,
        profile_source,
        emit_dispatch: Some(emitter::dispatch::dispatch),
        normalize_class: Some(normalize_class::normalize_class),
        register_tree: Some(tree_register::register_namespace_tree),
        expand_source: None,
    });
}

/// This crate as a [`vybe_runtime::Plugin`] — its `init` registers the
/// language with the shared framework.
pub struct Plugin;
impl vybe_runtime::Plugin for Plugin {
    fn name(&self) -> &'static str {
        "kotlin"
    }
    fn init(&self, _fw: &mut vybe_runtime::Framework<'_>) {
        register();
    }
}

// Link-time registration: this crate submits its plugin to the registry.
vybe_runtime::register_plugin!(Plugin);
