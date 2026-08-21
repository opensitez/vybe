//! WASM binary writer — Chunks → .wasm binary.
//!
//! Structured as a submodule:
//! - `types.rs` — GC type section (struct types from TypeEntry + array + func types)
//! - `sections.rs` — import, function, memory, export sections
//! - `code.rs` — code section (opcode translation)
//! - `proposals/` — one module per WebAssembly proposal: the imports it
//!   declares, the opcodes it emits, the custom sections it produces
//! - `builtins/` — the `wasm:js-*` import surface exposing JS-canonical
//!   collections (Array / Object / Map / Set / WeakMap / WeakSet /
//!   ArrayBuffer / DataView / 11 typed-arrays). Marshaling contract
//!   pinned in `JS_BUILTIN_CONVENTIONS.md`.

pub mod builtins;
pub mod code;
pub mod proposals;
pub mod sections;
pub mod types;

use proposals::{compilation_hints, exception_handling, extended_name_section, jspi};

use crate::encoding::*;
use vybe_runtime::Chunk;

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

    // Imported string constants (js-string-builtins § String constants) — one
    // imported global each. Collected before the globals so both the import
    // section and `global_map` index the same ordered list.
    let string_constants = sections::collect_string_constants(chunks);
    // Free globals — declared imports of the embedder, not module-owned state.
    let host_globals = sections::collect_host_globals(chunks);

    // The module-DEFINED globals, sliced from the compiler's index space.
    let globals = sections::collect_globals(chunks, &string_constants, &host_globals);

    // Type section: GC struct types + array type + function types
    let (type_section_data, type_ctx) =
        types::build_type_context(chunks, total_imports, &rt_imports);
    write_section(&mut out, SECTION_TYPE, &type_section_data);

    // Import section
    write_section(
        &mut out,
        SECTION_IMPORT,
        &sections::encode_import_section(
            chunks,
            &rt_imports,
            type_ctx.func_type_base,
            &string_constants,
            &host_globals,
        ),
    );

    // Function section
    write_section(
        &mut out,
        SECTION_FUNCTION,
        &sections::encode_func_section(chunks, total_imports, type_ctx.func_type_base),
    );

    // Table section — funcref table for call_indirect
    write_section(
        &mut out,
        4,
        &sections::encode_table_section(chunks, total_imports),
    );

    // Memory section
    let memory_section = if sections::module_uses_memory64(chunks) {
        sections::encode_memory64_section_with(1, None, false)
    } else {
        sections::encode_memory_section()
    };
    write_section(&mut out, SECTION_MEMORY, &memory_section);

    // Tag section (exception-handling + stack-switching proposals).
    // Always emits the `$vybe_exception (param externref)` tag when
    // the module uses `throw`, and additionally declares the
    // `$vybe_suspend (param externref) (result externref)` tag when
    // any `CONT_NEW` / `SUSPEND` / `RESUME` / `SWITCH` op appears.
    let emit_exception_tag = exception_handling::module_uses_exceptions(chunks);
    let suspend_idx = if type_ctx.uses_stack_switching {
        Some(type_ctx.suspend_tag_type_idx)
    } else {
        None
    };
    // The module's OWN tags follow the fixed prefix (exception tag, then the
    // stack-switching ones), so `VYBE_EXCEPTION_TAG` and the continuation tag
    // indices keep the values everything else already assumes.
    let reserved = 1 + u32::from(suspend_idx.is_some())
        + type_ctx.continuation_tag_type_indices.len() as u32;
    let tag_plan = exception_handling::plan_module_tags(chunks, reserved);
    if emit_exception_tag || type_ctx.uses_stack_switching {
        // Each declared tag needs a functype of its own arity; `(arity, 0)` is
        // exactly the blocktype shape `build_type_context` emitted for it.
        let extra_types: Vec<u32> = tag_plan
            .extra
            .iter()
            .map(|t| {
                type_ctx
                    .block_type_by_results
                    .get(&(t.arity, 0))
                    .copied()
                    .unwrap_or(type_ctx.exception_type_idx)
            })
            .collect();
        write_section(
            &mut out,
            SECTION_TAG,
            &exception_handling::encode_tag_section_full(
                type_ctx.exception_type_idx,
                suspend_idx,
                &type_ctx.continuation_tag_type_indices,
                &extra_types,
            ),
        );
    }

    // Global section — indexed externref globals
    if !globals.is_empty() {
        write_section(
            &mut out,
            SECTION_GLOBAL,
            &sections::encode_global_section(&globals),
        );
    }

    // Export section
    write_section(
        &mut out,
        SECTION_EXPORT,
        &sections::encode_export_section(chunks, total_imports),
    );

    // Element section — populate funcref table with chunk functions
    write_section(
        &mut out,
        9,
        &sections::encode_element_section(chunks, total_imports),
    );

    // branch_hint custom section — spec §branch-hinting requires this to
    // appear BEFORE the code section (not as a trailing custom section).
    if let Some(bh_payload) =
        compilation_hints::encode_branch_hint_payload(chunks, rt_imports.len())
    {
        let mut sec = Vec::new();
        write_name(&mut sec, compilation_hints::BRANCH_HINT_SECTION_NAME);
        sec.extend_from_slice(&bh_payload);
        write_section(&mut out, SECTION_CUSTOM, &sec);
    }

    // Code section
    write_section(
        &mut out,
        SECTION_CODE,
        &code::encode_code_section(chunks, &rt_imports, &type_ctx, &tag_plan),
    );

    // ── Trailing custom sections ─────────────────────────────────────
    // The standard `"name"` custom section (extended-name-section proposal) —
    // gives DevTools / profilers readable identifiers.
    let name_payload =
        extended_name_section::encode_name_section_payload(chunks, &rt_imports, &type_ctx);
    if !name_payload.is_empty() {
        let mut sec = Vec::new();
        write_name(&mut sec, "name");
        sec.extend_from_slice(&name_payload);
        write_section(&mut out, SECTION_CUSTOM, &sec);
    }

    // Compilation-hints proposal — tell the engine which functions to
    // optimize first. Skip the section when no hints apply.
    if let Some(co_payload) =
        compilation_hints::encode_compilation_order_payload(chunks, rt_imports.len())
    {
        let mut sec = Vec::new();
        write_name(&mut sec, compilation_hints::COMPILATION_ORDER_SECTION_NAME);
        sec.extend_from_slice(&co_payload);
        write_section(&mut out, SECTION_CUSTOM, &sec);
    }
    if let Some(in_payload) = compilation_hints::encode_inlining_payload(chunks, rt_imports.len()) {
        let mut sec = Vec::new();
        write_name(&mut sec, compilation_hints::INLINING_SECTION_NAME);
        sec.extend_from_slice(&in_payload);
        write_section(&mut out, SECTION_CUSTOM, &sec);
    }

    // JSPI custom section — `vybe.jspi` lists promising exports (wasm
    // function indices) that a JS host should wrap with
    // `WebAssembly.promising(...)` at load time so that Vybe's async
    // functions return real Promises across the JS boundary.
    if let Some(jspi_payload) = jspi::encode_payload(chunks, rt_imports.len()) {
        let mut sec = Vec::new();
        write_name(&mut sec, jspi::SECTION_NAME);
        sec.extend_from_slice(&jspi_payload);
        write_section(&mut out, SECTION_CUSTOM, &sec);
    }

    out
}

// ── Custom section (Vybe metadata for round-trip) ───────────────────────
// WASM Annotations Compliance (proposals/annotations/proposals/annotations/Overview.md):
// - Property descriptors are stored per-field using a flags byte (writable, enumerable, configurable)
// - Format: (@ecma262 descriptor field_name writable enumerable configurable)
// - Serialized in the "vybe" custom section with type entries
// - Follows WASM annotations proposal: metadata attached to struct fields in standardized format

fn encode_custom_section(chunks: &[Chunk]) -> Vec<u8> {
    let mut out = Vec::new();
    write_name(&mut out, "vybe");

    // Version
    out.push(3); // Version 3: adds exception TAG declarations (v2: type info + descriptors)

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

        // Type entries with WASM Annotations format field descriptors (v2+)
        write_leb128_u32(&mut out, chunk.types.len() as u32);
        for type_entry in &chunk.types {
            write_name(&mut out, &type_entry.name);
            // Supertype as the declared INDEX (0 = none), not a name.
            write_leb128_u32(&mut out, type_entry.parent_index as u32);

            // Fields with descriptors (WASM Annotations proposal @ecma262 namespace)
            write_leb128_u32(&mut out, type_entry.fields.len() as u32);
            for field in &type_entry.fields {
                write_name(&mut out, field);
                // Encode field descriptor as WASM annotation format
                if let Some(descriptor) = type_entry.field_descriptors.get(field) {
                    // Flags byte: bit 0 = writable, bit 1 = enumerable, bit 2 = configurable
                    let mut flags: u8 = 0;
                    if descriptor.writable {
                        flags |= 0x01;
                    }
                    if descriptor.enumerable {
                        flags |= 0x02;
                    }
                    if descriptor.configurable {
                        flags |= 0x04;
                    }
                    out.push(flags);
                } else {
                    // Standard descriptor: all flags set
                    out.push(0x07); // writable | enumerable | configurable
                }
            }

            // Methods
            write_leb128_u32(&mut out, type_entry.methods.len() as u32);
            for (method_name, chunk_idx) in &type_entry.methods {
                write_name(&mut out, method_name);
                write_leb128_u32(&mut out, *chunk_idx as u32);
            }

            // Other metadata
            out.push(if type_entry.is_interface { 1 } else { 0 });
            write_leb128_u32(&mut out, type_entry.implements.len() as u32);
            for &iface in &type_entry.implements {
                // Declared index, not a name — same as the supertype link.
                write_leb128_u32(&mut out, iface as u32);
            }
            if let Some(ctor_idx) = type_entry.constructor_chunk {
                out.push(1);
                write_leb128_u32(&mut out, ctor_idx as u32);
            } else {
                out.push(0);
            }
        }

        // Global imports (v3+): string constants are imported GLOBALS whose
        // key is interned in the constants table — the VM binds
        // `vm.globals[key]` from THIS list at instantiation. Without it, a
        // reloaded chunk's GLOBAL_GET finds nothing and every string
        // literal evaluates to undefined.
        write_leb128_u32(&mut out, chunk.global_imports.len() as u32);
        for gi in &chunk.global_imports {
            write_name(&mut out, &gi.module);
            write_name(&mut out, &gi.name);
        }

        // Exception tags (v3+). Without these, a reloaded module's
        // `throw`/`try_table` resolve tag indexes against an EMPTY map and
        // die at the first exception ("unknown exception tag index").
        // Imported-ness matters: imports resolve by name to a SHARED entity
        // at load (so `vybe:exception` is one entity across modules); locals
        // mint fresh entities per spec.
        write_leb128_u32(&mut out, chunk.tags.len() as u32);
        for tag in &chunk.tags {
            write_name(&mut out, &tag.debug_name);
            out.push(tag.arity);
            out.push(if tag.imported { 1 } else { 0 });
        }
    }
    out
}
