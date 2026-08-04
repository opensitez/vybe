//! Phase-0 resolution-snapshot harness (namespaceplan.md).
//!
//! For every language, run the ESM Link phase on an empty module (so the
//! bindings are exactly the profile's `esm_defaults` mounts) and diff the
//! resulting `(name → target)` pairs against the checked-in snapshot at
//! `tests/resolution_snapshots/<lang>.snap`.
//!
//! Self-initializing: a missing snapshot is written on first run and the
//! test passes. From then on, ANY change to what a name resolves to fails
//! this test — each namespace-unification phase must either keep the
//! snapshot byte-identical or update it as an explicit, reviewed diff.
//! This is the safety net against the `matchAll`-class silent hijacks.

use std::fs;
use std::path::PathBuf;

use vybe_compiler::ast::{Lang, Module};
use vybe_compiler::primitives::Compiler;

fn lang_enum(name: &str) -> Lang {
    match name {
        "vb" => Lang::VB,
        "js" => Lang::JavaScript,
        "csharp" => Lang::CSharp,
        "python" => Lang::Python,
        "ruby" => Lang::Ruby,
        "php" => Lang::PHP,
        "dart" => Lang::Dart,
        "pascal" => Lang::Pascal,
        "cobol" => Lang::Cobol,
        "fortran" => Lang::Fortran,
        "go" => Lang::Go,
        "lua" => Lang::Lua,
        "java" => Lang::Java,
        _ => Lang::Unknown, // c, wast — Link doesn't read the variant
    }
}

fn snapshot_dir() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("tests/resolution_snapshots")
}

#[test]
fn resolution_snapshots_are_stable() {
    let dir = snapshot_dir();
    fs::create_dir_all(&dir).expect("create snapshot dir");

    let mut failures: Vec<String> = Vec::new();
    let mut initialized: Vec<String> = Vec::new();

    for lang in vybe_compiler::languages::all() {
        let profile = match vybe_compiler::profile::parse_profile((lang.profile_source)()) {
            Ok(p) => p,
            Err(e) => {
                failures.push(format!("{}: profile parse failed: {e}", lang.name));
                continue;
            }
        };
        let module = Module {
            name: "snapshot".into(),
            language: lang_enum(lang.name),
            body: Vec::new(),
            imports: Vec::new(),
        scheduling: Default::default(),
        };
        let lines = Compiler::with_profile(profile).linked_resolution_snapshot(&module);
        let mut text = lines.join("\n");
        text.push('\n');

        let path = dir.join(format!("{}.snap", lang.name));
        if !path.exists() {
            fs::write(&path, &text).expect("write initial snapshot");
            initialized.push(lang.name.to_string());
            continue;
        }
        let expected = fs::read_to_string(&path).expect("read snapshot");
        if expected != text {
            // Line-level diff, kept short — the point is the signal.
            let mut diff = String::new();
            let old: Vec<&str> = expected.lines().collect();
            let new: Vec<&str> = text.lines().collect();
            for l in old.iter().filter(|l| !new.contains(l)).take(20) {
                diff.push_str(&format!("  - {l}\n"));
            }
            for l in new.iter().filter(|l| !old.contains(l)).take(20) {
                diff.push_str(&format!("  + {l}\n"));
            }
            failures.push(format!(
                "{}: resolution bindings changed (update {} ONLY as an \
                 intentional, reviewed diff):\n{diff}",
                lang.name,
                path.display()
            ));
        }
    }

    if !initialized.is_empty() {
        eprintln!(
            "resolution_snapshot: initialized snapshots for: {}",
            initialized.join(", ")
        );
    }
    assert!(
        failures.is_empty(),
        "resolution snapshot drift:\n{}",
        failures.join("\n")
    );
}
