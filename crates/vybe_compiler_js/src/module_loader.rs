//! Module loader — resolves, loads, and compiles JS modules.
//! Dependencies are concatenated into a single program before compilation.

use std::collections::HashSet;
use std::path::{Path, PathBuf};
use vybe_bytecode::Chunk;

use crate::Compiler;

/// Load a JS file and all its dependencies, compile as a single program.
/// Dependencies are prepended so their exports (globals) are available.
pub fn load_and_compile(entry_path: &Path) -> Result<Vec<Chunk>, String> {
    // Collect all source code in dependency order
    let mut loaded: HashSet<String> = HashSet::new();
    let mut combined_source = String::new();

    // Load dependencies first (recursive), then the entry file
    collect_sources(entry_path, &mut loaded, &mut combined_source)?;

    // Parse and compile the combined source as one program
    let program = vybe_parser_js::parse(&combined_source)
        .map_err(|e| format!("Parse error: {}", e))?;

    Compiler::new().compile(&program)
}

/// Recursively collect source code — dependencies first, then the file itself.
fn collect_sources(
    path: &Path,
    loaded: &mut HashSet<String>,
    output: &mut String,
) -> Result<(), String> {
    let canon = path.canonicalize()
        .unwrap_or_else(|_| path.to_path_buf())
        .to_string_lossy()
        .to_string();

    if loaded.contains(&canon) {
        return Ok(());
    }
    loaded.insert(canon);

    let source = std::fs::read_to_string(path)
        .map_err(|e| format!("Cannot read {}: {}", path.display(), e))?;

    // Parse to find imports
    let program = vybe_parser_js::parse(&source)
        .map_err(|e| format!("Parse error in {}: {}", path.display(), e))?;

    let base_dir = path.parent().unwrap_or(Path::new("."));

    // Recursively load dependencies first
    for stmt in &program.body {
        if let vybe_parser_js::Statement::Import { source: src, .. } = stmt {
            if src.starts_with("vybe:") || src.starts_with("js:") {
                continue; // host modules — resolved at runtime
            }
            let dep_path = resolve_module_path(base_dir, src)?;
            collect_sources(&dep_path, loaded, output)?;
        }
    }

    // Append this file's source (after dependencies)
    output.push_str("// --- module: ");
    output.push_str(&path.to_string_lossy());
    output.push_str(" ---\n");
    output.push_str(&source);
    output.push('\n');

    Ok(())
}

/// Resolve a module specifier to a file path.
fn resolve_module_path(base_dir: &Path, specifier: &str) -> Result<PathBuf, String> {
    let path = if specifier.starts_with("./") || specifier.starts_with("../") {
        base_dir.join(specifier)
    } else {
        base_dir.join(specifier)
    };

    // Add .js extension if missing
    let path = if path.extension().is_none() {
        path.with_extension("js")
    } else {
        path
    };

    if !path.exists() {
        return Err(format!("Module not found: {} (resolved to {})", specifier, path.display()));
    }

    Ok(path)
}
