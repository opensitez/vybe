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

use vybe_bytecode::Chunk;
use vybe_bytecode::component::{Component, Language, ExportImpl};
use std::collections::HashMap;

/// Build a Component from compiled chunks.
/// Scans chunks for top-level functions and classes, registers them as exports.
/// `name`: module name (filename without extension)
/// `language`: source language
/// `chunks`: compiled bytecode chunks (chunk 0 = script)
pub fn build_component(name: &str, language: Language, chunks: Vec<Chunk>) -> Component {
    let mut exports = HashMap::new();

    // Scan chunks for named functions (chunk 0 is script, others are functions/constructors)
    for (i, chunk) in chunks.iter().enumerate() {
        if i == 0 { continue; } // skip script chunk
        let fname = &chunk.name;
        if fname.is_empty() || fname.starts_with("__") || fname.starts_with("<") {
            continue; // skip internal/anonymous chunks
        }
        // Export as (module_name, function_name) → ChunkFn(i)
        exports.insert(
            (name.to_string(), fname.to_string()),
            ExportImpl::ChunkFn(i),
        );
    }

    // Also export type entries (classes/interfaces)
    let type_exports = HashMap::new(); // TODO: scan chunk[0].types

    Component {
        name: name.to_string(),
        language,
        imports: Vec::new(), // TODO: scan for unresolved global_gets
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
    component.imports.push((interface.to_string(), func_name.to_string()));
}
