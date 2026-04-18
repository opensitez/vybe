//! WASM binary format reader/writer.
//!
//! Structured as a submodule:
//! - `encoding.rs` — constants, LEB128, section helpers
//! - `types.rs` — GC type section (struct types from TypeEntry + array + func types)
//! - `sections.rs` — import, function, memory, export sections
//! - `code.rs` — code section (opcode translation)
//! - `reader.rs` — .wasm binary reader

pub mod encoding;
pub mod types;
pub mod sections;
pub mod code;
pub mod reader;

use encoding::*;
use crate::Chunk;

// ── Writer ──────────────────────────────────────────────────────────────

pub fn write_wasm(chunks: &[Chunk]) -> Vec<u8> {
    let mut out = Vec::new();
    out.extend_from_slice(&WASM_MAGIC);
    out.extend_from_slice(&WASM_VERSION);

    // Custom section: Vybe metadata for round-trip
    write_section(&mut out, SECTION_CUSTOM, &encode_custom_section(chunks));

    // Collect imports — total_imports = host imports + wasm:js-* builtins
    let rt_imports = sections::collect_rt_imports(chunks);
    let host_import_count = chunks.first().map(|c| c.imports.len()).unwrap_or(0);
    let total_imports = host_import_count + rt_imports.len(); // ALL WASM-level imports

    // Collect globals (string-keyed → indexed)
    let (globals, global_map) = sections::collect_globals(chunks);

    // Type section: GC struct types + array type + function types
    let (type_section_data, type_ctx) = types::build_type_context(chunks, total_imports, &rt_imports);
    write_section(&mut out, SECTION_TYPE, &type_section_data);

    // Import section
    write_section(&mut out, SECTION_IMPORT, &sections::encode_import_section(chunks, &rt_imports, type_ctx.func_type_base));

    // Function section
    write_section(&mut out, SECTION_FUNCTION, &sections::encode_func_section(chunks, total_imports, type_ctx.func_type_base));

    // Table section — funcref table for call_indirect
    write_section(&mut out, 4, &sections::encode_table_section(chunks, total_imports));

    // Memory section
    write_section(&mut out, SECTION_MEMORY, &sections::encode_memory_section());

    // Global section — indexed externref globals
    if !globals.is_empty() {
        write_section(&mut out, SECTION_GLOBAL, &sections::encode_global_section(&globals));
    }

    // Export section
    write_section(&mut out, SECTION_EXPORT, &sections::encode_export_section(chunks, total_imports));

    // Element section — populate funcref table with chunk functions
    write_section(&mut out, 9, &sections::encode_element_section(chunks, total_imports));

    // Code section
    write_section(&mut out, SECTION_CODE, &code::encode_code_section(chunks, &rt_imports, &type_ctx, &global_map));

    out
}

// ── Reader ──────────────────────────────────────────────────────────────

pub fn read_wasm(data: &[u8]) -> Result<Vec<Chunk>, String> {
    reader::read_wasm(data)
}

// ── Custom section (Vybe metadata for round-trip) ───────────────────────

fn encode_custom_section(chunks: &[Chunk]) -> Vec<u8> {
    let mut out = Vec::new();
    write_name(&mut out, "vybe");

    // Version
    out.push(1);

    // Number of chunks
    write_leb128_u32(&mut out, chunks.len() as u32);

    for chunk in chunks {
        // Chunk metadata
        write_name(&mut out, &chunk.name);
        out.push(chunk.arity);
        write_leb128_u32(&mut out, chunk.local_count as u32);

        // Constants
        write_leb128_u32(&mut out, chunk.constants.len() as u32);
        for c in &chunk.constants {
            encode_value(&mut out, c);
        }

        // Imports (only on chunk 0)
        write_leb128_u32(&mut out, chunk.imports.len() as u32);
        for import in &chunk.imports {
            write_name(&mut out, &import.module);
            write_name(&mut out, &import.name);
        }

        // Bytecode
        write_leb128_u32(&mut out, chunk.code.len() as u32);
        out.extend_from_slice(&chunk.code);

        // Line info
        write_leb128_u32(&mut out, chunk.lines.len() as u32);
        for &line in &chunk.lines {
            write_leb128_u32(&mut out, line);
        }
    }
    out
}
