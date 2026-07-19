//! Component Model helpers — build cross-language module metadata.
//!
//! After compiling source code to chunks, each compiler calls these helpers
//! to produce a Component with proper export declarations. This enables the
//! Linker to resolve cross-language imports.
//!
//! Example: Python defines `def greet(name):`, Dart calls `greet("world")`.
//! The Python compiler emits a Component with export ("module", "greet").
//! The Dart compiler emits a Component with import ("module", "greet").
//! The Linker wires them together.

use std::collections::HashMap;
use vybe_bytecode::component::{Component, ExportImpl, Language};
use vybe_bytecode::Chunk;

/// Build a Component from compiled chunks.
/// Scans chunks for top-level functions and classes, registers them as exports.
/// `name`: module name (filename without extension)
/// `language`: source language
/// `chunks`: compiled bytecode chunks (chunk 0 = script)
pub fn build_component(name: &str, language: Language, chunks: Vec<Chunk>) -> Component {
    let mut exports = HashMap::new();

    // Scan chunks for named functions (chunk 0 is script, others are functions/constructors)
    for (i, chunk) in chunks.iter().enumerate() {
        if i == 0 {
            continue;
        } // skip script chunk
        let fname = &chunk.name;
        if fname.is_empty() || fname.starts_with("__") || fname.starts_with("<") {
            continue; // skip internal/anonymous chunks
        }
        // Export under both the module-specific key AND the wildcard key.
        // Module-specific: ("mod_name", "square") — for explicit module imports
        // Wildcard: ("*", "square") — for cross-language resolution
        exports.insert(
            (name.to_string(), fname.to_string()),
            ExportImpl::ChunkFn(i),
        );
        exports.insert(("*".to_string(), fname.to_string()), ExportImpl::ChunkFn(i));
        // Also export lowercase for case-insensitive languages
        let lower = fname.to_lowercase();
        if lower != *fname {
            exports.insert(("*".to_string(), lower.clone()), ExportImpl::ChunkFn(i));
            exports.insert((name.to_string(), lower), ExportImpl::ChunkFn(i));
        }
    }

    // Export class constructors from the type table
    if !chunks.is_empty() {
        for entry in &chunks[0].types {
            if let Some(ctor_idx) = entry.constructor_chunk {
                exports.insert(
                    ("*".to_string(), entry.name.clone()),
                    ExportImpl::ChunkFn(ctor_idx),
                );
                exports.insert(
                    (name.to_string(), entry.name.clone()),
                    ExportImpl::ChunkFn(ctor_idx),
                );
            }
        }
    }

    // Collect imports from the script chunk's import table
    // (compilers emit call_import("*", name) for unresolved references)
    let mut imports = Vec::new();
    if !chunks.is_empty() {
        for imp in &chunks[0].imports {
            imports.push((imp.module.clone(), imp.name.clone()));
        }
    }

    let type_exports = HashMap::new();

    Component {
        name: name.to_string(),
        language,
        imports,
        exports,
        chunks,
        type_exports,
        type_imports: Vec::new(),
    }
}

/// Register a specific function as an export.
/// Use this when the compiler knows exactly which functions should be public.
pub fn add_export(component: &mut Component, interface: &str, func_name: &str, chunk_idx: usize) {
    component.exports.insert(
        (interface.to_string(), func_name.to_string()),
        ExportImpl::ChunkFn(chunk_idx),
    );
}

/// Register an import requirement.
/// Use this when the compiler encounters an unresolved reference to another module.
pub fn add_import(component: &mut Component, interface: &str, func_name: &str) {
    component
        .imports
        .push((interface.to_string(), func_name.to_string()));
}
