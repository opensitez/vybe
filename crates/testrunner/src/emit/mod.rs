//! Per-language emitters.
//!
//! An emitter only knows how to rewrite a case's output calls into assertions
//! against the language's harness. The harness itself is never here — it is
//! real source under `harness/<lang>/`, so it can be read, formatted and
//! debugged with that language's own tools.

pub mod c;
pub mod cobol;
pub mod csharp;
pub mod dart;
pub mod fortran;
pub mod go;
pub mod java;
pub mod js;
pub mod kotlin;
pub mod lua;
pub mod pascal;
pub mod php;
pub mod python;
pub mod ruby;
pub mod vb;
pub mod wast;

use std::path::PathBuf;

/// Where the harness sources live — real files in the language under test.
pub fn harness_path(lang: &str) -> Option<PathBuf> {
    let ext = extension(lang)?;
    let path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("harness")
        .join(lang)
        .join(format!("check.{ext}"));
    path.exists().then_some(path)
}

/// The harness stripped of the preamble a test file already carries.
///
/// `harness/<lang>/check.<ext>` stays a COMPLETE, valid source file — that is
/// the whole point of it being Go rather than a Rust string; it has to open in
/// an editor, run through `gofmt`, and be debuggable on its own. Splicing it
/// into a test therefore has to drop the parts that would be declared twice:
/// the package clause, the imports, and the file header comment.
///
/// Inlining rather than passing the harness as a second file is forced, not
/// preferred: `vybex a.go b.go` concatenates its sources, so a second
/// `package main` is a parse error, and `go run` refuses files from two
/// different directories. Both are recorded as findings.
pub fn harness_body(lang: &str) -> anyhow::Result<String> {
    let path = harness_path(lang).ok_or_else(|| anyhow::anyhow!("no harness for `{lang}`"))?;
    let text = std::fs::read_to_string(&path)?;

    let mut body = String::new();
    let mut in_import_block = false;
    for line in text.lines() {
        let t = line.trim();
        if in_import_block {
            in_import_block = t != ")";
            continue;
        }
        if t.starts_with("import (") {
            in_import_block = true;
            continue;
        }
        if t.starts_with("package ") || t.starts_with("import ") {
            continue;
        }
        if t.starts_with("use ") || t.starts_with("#!") || t == "<?php" {
            continue;
        }
        // The file header explains the harness; a test file wants the code.
        if body.is_empty() && (t.is_empty() || t.starts_with("//")) {
            continue;
        }
        body.push_str(line);
        body.push('\n');
    }
    Ok(body.trim_end().to_string())
}

/// Source-file extension for a language directory name.
pub fn extension(lang: &str) -> Option<&'static str> {
    Some(match lang {
        "go" => "go",
        "python" => "py",
        "js" => "js",
        "php" => "php",
        "lua" => "lua",
        "ruby" => "rb",
        "dart" => "dart",
        "kotlin" => "kt",
        "java" => "java",
        "csharp" => "cs",
        "c" => "c",
        "pascal" => "pas",
        "fortran" => "f90",
        "cobol" => "cob",
        "vb" => "vb",
        "wast" => "wat",
        _ => return None,
    })
}
