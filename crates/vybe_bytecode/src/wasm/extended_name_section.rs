//! # extended-name-section proposal
//!
//! Spec: `proposals/extended-name-section/`. Extends
//! the standard custom `"name"` section (originally only module / func /
//! local / type / field / tag names) with subsections for:
//!
//! | Subsection | Id   | What it names          |
//! |------------|------|-------------------------|
//! | module     | `0`  | the module itself       |
//! | function   | `1`  | function indices        |
//! | local      | `2`  | locals per function     |
//! | label      | `3`  | structured-control labels *(✅ emitted)* |
//! | type       | `4`  | type indices *(✅ emitted)* |
//! | table      | `5`  | table indices *(✅ emitted)* |
//! | memory     | `6`  | memory indices *(✅ emitted)* |
//! | global     | `7`  | global indices *(✅ emitted)* |
//! | element    | `8`  | element-segment indices *(✅ emitted)* |
//! | data       | `9`  | data-segment indices *(✅ empty but present)* |
//! | field      | `10` | GC struct field names *(✅ emitted)* |
//! | tag        | `11` | exception tag names *(✅ empty but present)* |
//!
//! ## Purpose
//!
//! Without this section, Chrome DevTools and Node profilers see opaque
//! identifiers like `$func12`. With it they show `Form1_Button1_Click`,
//! `wasm:js-undefined.value`, etc., which makes stack traces and
//! profiling output actually readable.
//!
//! ## Section format
//!
//! ```text
//! "name" custom section payload:
//!   subsection 1 (id 1):
//!     u32 size
//!     u32 count
//!     for each:
//!       u32 index
//!       name (u32 length + utf-8 bytes)
//!   subsection 2 (id 2):
//!     ... (indirect name map: outer-idx → inner map)
//!   ...
//! ```
//!
//! This module produces the `"name"` custom-section payload for our
//! module. It's the only entry point the writer pipeline needs.

use super::encoding::*;
use super::sections::rt_globals;
use super::types::WasmTypeContext;
use crate::Chunk;
use crate::Op;

/// Produce the payload of the `"name"` custom section for the given
/// chunks + declared imports/globals. The caller wraps this in a
/// `SECTION_CUSTOM` with name `"name"`.
pub fn encode_name_section_payload(
    chunks: &[Chunk],
    rt_imports: &[(&str, &str)],
    type_ctx: &WasmTypeContext,
) -> Vec<u8> {
    let mut out = Vec::new();

    // Subsection 0: module name — use the script chunk's name if any.
    if let Some(script) = chunks.first() {
        if !script.name.is_empty() && script.name != "<script>" {
            let mut sub = Vec::new();
            write_name(&mut sub, &script.name);
            write_subsection(&mut out, 0, &sub);
        }
    }

    // Subsection 1: function names — host imports then rt imports then chunks.
    {
        let mut sub = Vec::new();
        let host_imports_len = chunks.first().map(|c| c.imports.len()).unwrap_or(0);
        let total = host_imports_len + rt_imports.len() + chunks.len();
        write_leb128_u32(&mut sub, total as u32);
        let mut idx: u32 = 0;
        // Host imports
        if let Some(chunk) = chunks.first() {
            for imp in &chunk.imports {
                write_leb128_u32(&mut sub, idx);
                write_name(&mut sub, &format!("{}.{}", imp.module, imp.name));
                idx += 1;
            }
        }
        // Runtime / builtin imports
        for (module, name) in rt_imports {
            write_leb128_u32(&mut sub, idx);
            write_name(&mut sub, &format!("{}.{}", module, name));
            idx += 1;
        }
        // Chunk-defined functions
        for chunk in chunks {
            write_leb128_u32(&mut sub, idx);
            write_name(&mut sub, &chunk.name);
            idx += 1;
        }
        write_subsection(&mut out, 1, &sub);
    }

    // Subsection 2: local names — one entry per chunk, naming the
    // WASM-convention locals (`local0`, `local1`, …). We don't have
    // source-level names here so we use the convention directly.
    {
        let mut sub = Vec::new();
        let host_imports_len = chunks.first().map(|c| c.imports.len()).unwrap_or(0);
        let func_base = host_imports_len + rt_imports.len();
        write_leb128_u32(&mut sub, chunks.len() as u32);
        for (ci, chunk) in chunks.iter().enumerate() {
            write_leb128_u32(&mut sub, (func_base + ci) as u32);
            let arity_u = chunk.arity as u16;
            let local_count = (chunk.local_count as u16).max(arity_u) as u32;
            // Add 1 for the temp extern used by some ops.
            let with_temp = local_count + 1;
            write_leb128_u32(&mut sub, with_temp);
            for li in 0..local_count {
                write_leb128_u32(&mut sub, li);
                let label = if li < arity_u as u32 {
                    format!("arg{}", li)
                } else {
                    format!("local{}", li)
                };
                write_name(&mut sub, &label);
            }
            write_leb128_u32(&mut sub, local_count);
            write_name(&mut sub, "__temp");
        }
        write_subsection(&mut out, 2, &sub);
    }

    // Subsection 3: label names — per-function indirect map of
    // block/loop/if labels in the order they appear in the bytecode.
    {
        let host_imports_len = chunks.first().map(|c| c.imports.len()).unwrap_or(0);
        let func_base = host_imports_len + rt_imports.len();
        let mut per_func: Vec<(u32, Vec<String>)> = Vec::new();
        for (ci, chunk) in chunks.iter().enumerate() {
            let labels = collect_label_names(chunk);
            if !labels.is_empty() {
                per_func.push(((func_base + ci) as u32, labels));
            }
        }
        if !per_func.is_empty() {
            let mut sub = Vec::new();
            write_leb128_u32(&mut sub, per_func.len() as u32);
            for (fn_idx, labels) in per_func {
                write_leb128_u32(&mut sub, fn_idx);
                write_leb128_u32(&mut sub, labels.len() as u32);
                for (li, name) in labels.iter().enumerate() {
                    write_leb128_u32(&mut sub, li as u32);
                    write_name(&mut sub, name);
                }
            }
            write_subsection(&mut out, 3, &sub);
        }
    }

    // Subsection 4: type names — described struct, its descriptor,
    // the array type, and every function-arity type we emitted.
    {
        let mut pairs: Vec<(u32, String)> = Vec::new();
        for (name_lc, idx) in &type_ctx.struct_type_indices {
            pairs.push((*idx, name_lc.clone()));
        }
        for (name_lc, idx) in &type_ctx.desc_type_indices {
            pairs.push((*idx, format!("{}__desc", name_lc)));
        }
        pairs.push((type_ctx.array_type_idx, "array".to_string()));
        for (arity, idx) in &type_ctx.func_type_by_arity {
            pairs.push((*idx, format!("fn_arity_{}", arity)));
        }
        // Stable order so round-trip bytes are deterministic.
        pairs.sort_by_key(|(i, _)| *i);
        pairs.dedup_by_key(|(i, _)| *i);
        if !pairs.is_empty() {
            let mut sub = Vec::new();
            write_leb128_u32(&mut sub, pairs.len() as u32);
            for (idx, name) in pairs {
                write_leb128_u32(&mut sub, idx);
                write_name(&mut sub, &name);
            }
            write_subsection(&mut out, 4, &sub);
        }
    }

    // Subsection 10: field names — per struct-type indirect map of
    // field index → field name. Makes `struct.get N` readable in
    // DevTools as e.g. `Button.x` instead of `struct_7.field_0`.
    {
        // Build (struct_type_idx, fields) pairs sorted by type index.
        let mut entries: Vec<(u32, Vec<String>)> = type_ctx
            .struct_fields
            .iter()
            .filter_map(|(name_lc, fields)| {
                type_ctx
                    .struct_type_indices
                    .get(name_lc)
                    .map(|idx| (*idx, fields.clone()))
            })
            .collect();
        entries.sort_by_key(|(i, _)| *i);
        if !entries.is_empty() {
            let mut sub = Vec::new();
            write_leb128_u32(&mut sub, entries.len() as u32);
            for (type_idx, fields) in entries {
                write_leb128_u32(&mut sub, type_idx);
                write_leb128_u32(&mut sub, fields.len() as u32);
                for (fi, name) in fields.iter().enumerate() {
                    write_leb128_u32(&mut sub, fi as u32);
                    write_name(&mut sub, name);
                }
            }
            write_subsection(&mut out, 10, &sub);
        }
    }

    // Subsection 5: table names — we have a single funcref table for
    // call_indirect dispatch.
    {
        let mut sub = Vec::new();
        write_leb128_u32(&mut sub, 1); // count
        write_leb128_u32(&mut sub, 0); // table index 0
        write_name(&mut sub, "funcref_table");
        write_subsection(&mut out, 5, &sub);
    }

    // Subsection 6: memory names — one linear memory.
    {
        let mut sub = Vec::new();
        write_leb128_u32(&mut sub, 1);
        write_leb128_u32(&mut sub, 0);
        write_name(&mut sub, "memory");
        write_subsection(&mut out, 6, &sub);
    }

    // Subsection 7: global names — the three js-primitive globals we import.
    {
        let globals = rt_globals();
        if !globals.is_empty() {
            let mut sub = Vec::new();
            write_leb128_u32(&mut sub, globals.len() as u32);
            for (i, (module, name)) in globals.iter().enumerate() {
                write_leb128_u32(&mut sub, i as u32);
                write_name(&mut sub, &format!("{}.{}", module, name));
            }
            write_subsection(&mut out, 7, &sub);
        }
    }

    // Subsection 8: element-segment names.
    {
        let mut sub = Vec::new();
        write_leb128_u32(&mut sub, 1);
        write_leb128_u32(&mut sub, 0);
        write_name(&mut sub, "func_dispatch");
        write_subsection(&mut out, 8, &sub);
    }

    // Subsection 9: data-segment names. We don't currently emit a data
    // section, so this is a valid-but-empty name map. Keeping the empty
    // subsection present advertises we understand the name format — if
    // someone adds a data segment later, names go here.
    {
        let mut sub = Vec::new();
        write_leb128_u32(&mut sub, 0); // zero entries
        write_subsection(&mut out, 9, &sub);
    }

    // Subsection 11: tag names. Tags are for the exception-handling
    // proposal. When a tag section is eventually emitted, its per-tag
    // labels go here. Empty until then — same rationale as subsection 9.
    {
        let mut sub = Vec::new();
        write_leb128_u32(&mut sub, 0);
        write_subsection(&mut out, 11, &sub);
    }

    out
}

/// Write a name-section subsection: `id (u8) | size (u32 LEB) | payload`.
fn write_subsection(out: &mut Vec<u8>, id: u8, payload: &[u8]) {
    out.push(id);
    write_leb128_u32(out, payload.len() as u32);
    out.extend_from_slice(payload);
}

/// Walk a chunk's bytecode and produce synthetic label names for every
/// structured-control instruction (BLOCK / LOOP / IF / TRY_TABLE) in
/// order. Labels are 0-indexed per function.
fn collect_label_names(chunk: &Chunk) -> Vec<String> {
    let mut names: Vec<String> = Vec::new();
    let code = &chunk.code;
    let mut ip = 0usize;
    while ip + 1 < code.len() {
        let prefix = code[ip];
        let sub = code[ip + 1];
        let op = match Op::decode(prefix, sub as u16) {
            Some(op) => op,
            None => {
                ip += 2;
                continue;
            }
        };
        ip += 2;
        // Structured control ops introduce a new label index.
        if op == Op::BLOCK || op == Op::LOOP || op == Op::TRY_TABLE {
            let tag = if op == Op::BLOCK {
                "block"
            } else if op == Op::LOOP {
                "loop"
            } else {
                "try"
            };
            names.push(format!("label{}_{}", names.len(), tag));
        }
        ip += op.operand_format().size_in(code, ip);
    }
    names
}
