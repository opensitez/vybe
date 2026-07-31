//! Linking several independent source files into one bundle.
//!
//! Sources are concatenated into a single string and parsed once
//! (`Bundle::prepared_module`), which is what gives cross-file declaration
//! visibility — the C-linker model the multi-file CLI relies on.
//!
//! KNOWN GAP (not covered here because it is unfixed): that same concatenation
//! resolves EVERY file's relative imports against the FIRST source's directory
//! (`bundle.rs`, "Resolve imports relative to first source's directory"), so a
//! second file importing its own sibling looks in the wrong folder and the
//! import is skipped with a warning. Parsing per source is NOT the fix — the
//! language `parse` hook injects the JS prelude, so N sources would emit N
//! copies of it. Fixing it needs either a fragment-parse hook (parse a source
//! with no prelude) or per-source import-specifier normalization before
//! concatenation.

use std::path::PathBuf;

use vybe_compiler::bundle::{Bundle, EntryPoint, SourceFile};

fn js_language() -> vybe_compiler::languages::Language {
    static R: std::sync::Once = std::sync::Once::new();
    R.call_once(vybe_language_js::register);
    vybe_compiler::languages::find_by_name("js").expect("js language registered")
}

fn bundle(sources: Vec<(&str, &str)>) -> Bundle {
    Bundle {
        name: "multi".to_string(),
        language: js_language(),
        sources: sources
            .into_iter()
            .map(|(path, code)| SourceFile {
                path: PathBuf::from(path),
                code: code.to_string(),
            })
            .collect(),
        wasm_files: Vec::new(),
        entry_point: EntryPoint::Auto,
    }
}

/// Declarations cross file boundaries: every source's top-level declarations
/// land in the one merged module, in source order.
#[test]
fn declarations_from_every_source_reach_the_merged_module() {
    let module = bundle(vec![
        ("/proj/a.js", "function fromA() { return 1; }\n"),
        ("/proj/b.js", "function fromB() { return 2; }\n"),
    ])
    .prepared_module()
    .expect("prepared module");

    let names: Vec<String> = module
        .body
        .iter()
        .filter_map(|stmt| match &stmt.kind {
            vybe_ast::StmtKind::FunctionDecl { name, .. } => Some(name.clone()),
            _ => None,
        })
        .collect();
    let a_at = names.iter().position(|n| n == "fromA");
    let b_at = names.iter().position(|n| n == "fromB");
    assert!(
        a_at.is_some() && b_at.is_some(),
        "both files' declarations must be present, got {names:?}"
    );
    assert!(
        a_at < b_at,
        "sources must merge in order (a before b), got {names:?}"
    );
}

/// A single source is unaffected by any multi-source handling.
#[test]
fn single_source_reaches_the_module_body() {
    let module = bundle(vec![("/proj/only.js", "let onlyFileMarker = 41;\n")])
        .prepared_module()
        .expect("prepared module");
    let has_decl = module.body.iter().any(|stmt| match &stmt.kind {
        vybe_ast::StmtKind::VarDecl { declarations, .. } => declarations.iter().any(|d| {
            matches!(&d.pattern, vybe_ast::BindingPattern::Ident(n) if n == "onlyFileMarker")
        }),
        _ => false,
    });
    assert!(has_decl, "single source must still reach the module body");
}
