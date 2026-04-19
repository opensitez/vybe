//! WASM type section encoding with Custom Descriptors.
//!
//! Each TypeEntry from the compiler produces TWO WASM GC types:
//! 1. The described struct type (object fields as externref)
//! 2. The descriptor struct type (JS prototype + vtable methods)
//!
//! This follows the Custom Descriptors proposal:
//! proposals/custom-descriptors/proposals/custom-descriptors/Overview.md

use super::encoding::*;
use crate::Chunk;

// Custom Descriptors binary encoding
const CD_DESCRIPTOR: u8 = 0x4D;  // (descriptor $x) prefix
const CD_DESCRIBES: u8 = 0x4C;   // (describes $x) prefix
const CD_SUB_FINAL: u8 = 0x4F;   // sub final

/// Context for .wasm emission — maps internal types to WASM type indices.
pub struct WasmTypeContext {
    /// type_name (lowercased) → WASM type index for the described struct
    pub struct_type_indices: std::collections::HashMap<String, u32>,
    /// type_name → WASM type index for the descriptor struct (vtable + proto)
    pub desc_type_indices: std::collections::HashMap<String, u32>,
    /// type_name → vec of field names in order (for field index lookup)
    pub struct_fields: std::collections::HashMap<String, Vec<String>>,
    /// WASM type index for the dynamic array type
    pub array_type_idx: u32,
    /// First function type index (after GC types)
    pub func_type_base: u32,
    /// Total number of GC types (structs + descriptors + array)
    pub gc_type_count: u32,
    /// arity → WASM type index for (externref * arity) -> externref
    pub func_type_by_arity: std::collections::HashMap<u8, u32>,
}

impl WasmTypeContext {
    /// Look up the WASM type index for a described struct type by name.
    pub fn struct_type(&self, name: &str) -> Option<u32> {
        self.struct_type_indices.get(&name.to_lowercase()).copied()
    }

    /// Look up the WASM type index for a descriptor type by name.
    pub fn desc_type(&self, name: &str) -> Option<u32> {
        self.desc_type_indices.get(&name.to_lowercase()).copied()
    }

    /// Look up the field index for a field name within a struct type.
    pub fn field_index(&self, type_name: &str, field_name: &str) -> Option<u32> {
        let fields = self.struct_fields.get(&type_name.to_lowercase())?;
        fields.iter().position(|f| f == &field_name.to_lowercase()).map(|i| i as u32)
    }
}

/// Build the type context and encode the type section.
/// Layout: [rec group: (described struct + descriptor struct) per TypeEntry] [array type] [function types]
pub fn build_type_context(chunks: &[Chunk], import_count: usize, rt_imports: &[(&str, &str)]) -> (Vec<u8>, WasmTypeContext) {
    let mut out = Vec::new();
    let mut ctx = WasmTypeContext {
        struct_type_indices: std::collections::HashMap::new(),
        desc_type_indices: std::collections::HashMap::new(),
        struct_fields: std::collections::HashMap::new(),
        array_type_idx: 0,
        func_type_base: 0,
        gc_type_count: 0,
        func_type_by_arity: std::collections::HashMap::new(),
    };

    // Collect TypeEntry definitions from chunk 0
    let type_entries: Vec<&crate::chunk::TypeEntry> = chunks.first()
        .map(|c| c.types.iter().collect())
        .unwrap_or_default();

    // Layout:
    // For each TypeEntry: 2 types (described struct + descriptor struct)
    // Then: 1 array type
    // Then: function types for imports + chunks
    let gc_struct_pairs = type_entries.len() as u32;
    let array_count = 1u32;
    let func_count = (import_count + chunks.len()) as u32;
    // Each TypeEntry produces 2 types in a rec group
    let gc_type_count = gc_struct_pairs * 2 + array_count;
    ctx.gc_type_count = gc_type_count;
    ctx.array_type_idx = gc_struct_pairs * 2;
    ctx.func_type_base = gc_type_count;

    let total = gc_type_count + func_count;
    write_leb128_u32(&mut out, total);

    // ── GC struct types with custom descriptors ──
    // Each TypeEntry becomes a rec group of (described, descriptor)
    for (i, te) in type_entries.iter().enumerate() {
        let described_idx = (i as u32) * 2;
        let descriptor_idx = (i as u32) * 2 + 1;
        let name_lower = te.name.to_lowercase();
        ctx.struct_type_indices.insert(name_lower.clone(), described_idx);
        ctx.desc_type_indices.insert(name_lower.clone(), descriptor_idx);
        ctx.struct_fields.insert(name_lower, te.fields.clone());

        // Described struct: (descriptor $desc_idx) (struct (field (mut externref))*)
        // Binary: CD_SUB_FINAL 0_supertypes CD_DESCRIPTOR desc_idx GC_STRUCT field_count fields...
        out.push(CD_SUB_FINAL);
        write_leb128_u32(&mut out, 0); // 0 supertypes (TODO: use te.parent)
        out.push(CD_DESCRIPTOR);
        write_leb128_u32(&mut out, descriptor_idx);
        out.push(GC_STRUCT);
        write_leb128_u32(&mut out, te.fields.len() as u32);
        for _ in &te.fields {
            out.push(TYPE_EXTERNREF);
            out.push(GC_MUT);
        }

        // Descriptor struct: (describes $described_idx) (struct (field $proto externref) (field $method funcref)*)
        // Binary: CD_SUB_FINAL 0_supertypes CD_DESCRIBES described_idx GC_STRUCT field_count fields...
        out.push(CD_SUB_FINAL);
        write_leb128_u32(&mut out, 0);
        out.push(CD_DESCRIBES);
        write_leb128_u32(&mut out, described_idx);
        out.push(GC_STRUCT);
        let desc_field_count = 1 + te.methods.len(); // proto + methods
        write_leb128_u32(&mut out, desc_field_count as u32);
        // First field: JS prototype (externref, immutable)
        out.push(TYPE_EXTERNREF);
        out.push(GC_IMMUT);
        // Method fields: funcref for each method
        for _ in &te.methods {
            out.push(TYPE_FUNCREF);
            out.push(GC_IMMUT);
        }
    }

    // ── Array type ──
    // (array (mut externref)) — for dynamic arrays
    out.push(GC_ARRAY);
    out.push(TYPE_EXTERNREF);
    out.push(GC_MUT);

    // ── Function types ──
    // ── Function types with proper signatures ──
    // Each unique import signature needs its own type.
    // wasm:js-number builtins have typed params (i32, f64) and externref results.
    // Dynamic language imports use externref for everything.
    // Chunk functions use externref params/results.

    // For now, use distinct types per arity.
    // Type for 0-param imports: () -> externref
    // Type for 1-param (externref) imports: (externref) -> externref
    // Type for chunk functions: (externref * arity) -> externref
    // TODO: wasm:js-number needs (i32)->externref, (f64)->externref, (externref)->f64 etc.

    // Import function types — per-import typed signatures
    // Host imports from chunk 0 — scan CALL_IMPORT bytecode to find actual arity
    let host_import_count = chunks.first().map(|c| c.imports.len()).unwrap_or(0);
    let mut host_arity: Vec<u8> = vec![0; host_import_count];
    for chunk in chunks {
        let mut ip = 0;
        while ip < chunk.code.len() {
            if ip + 1 >= chunk.code.len() { break; }
            if let Some(op) = crate::opcode::Op::decode(chunk.code[ip], chunk.code[ip + 1]) {
                if op == crate::opcode::Op::CALL_IMPORT {
                    let import_idx = ((chunk.code[ip + 2] as u16) << 8) | chunk.code[ip + 3] as u16;
                    let argc = chunk.code[ip + 4];
                    if (import_idx as usize) < host_import_count {
                        host_arity[import_idx as usize] = host_arity[import_idx as usize].max(argc);
                    }
                }
                ip += super::code::opcode_size(op, &chunk.code, ip);
            } else {
                ip += 2;
            }
        }
    }
    for i in 0..host_import_count {
        let argc = host_arity[i];
        out.push(TYPE_FUNC);
        write_leb128_u32(&mut out, argc as u32);
        for _ in 0..argc { out.push(TYPE_EXTERNREF); }
        write_leb128_u32(&mut out, 1);
        out.push(TYPE_EXTERNREF);
    }
    // Runtime imports: typed signatures based on module:name
    for &(module, name) in rt_imports {
        out.push(TYPE_FUNC);
        match (module, name) {
            ("wasm:js-number", "fromI32") => {
                write_leb128_u32(&mut out, 1); // 1 param
                out.push(TYPE_I32);
                write_leb128_u32(&mut out, 1); // 1 result
                out.push(TYPE_EXTERNREF);
            }
            ("wasm:js-number", "fromF64") => {
                write_leb128_u32(&mut out, 1);
                out.push(TYPE_F64);
                write_leb128_u32(&mut out, 1);
                out.push(TYPE_EXTERNREF);
            }
            ("wasm:js-number", "toF64") => {
                write_leb128_u32(&mut out, 1);
                out.push(TYPE_EXTERNREF);
                write_leb128_u32(&mut out, 1);
                out.push(TYPE_F64);
            }
            ("wasm:js-number", "toI32") => {
                write_leb128_u32(&mut out, 1);
                out.push(TYPE_EXTERNREF);
                write_leb128_u32(&mut out, 1);
                out.push(TYPE_I32);
            }
            ("wasm:js-string", "test") | ("wasm:js-boolean", "test") | ("wasm:js-undefined", "test") => {
                write_leb128_u32(&mut out, 1);
                out.push(TYPE_EXTERNREF);
                write_leb128_u32(&mut out, 1);
                out.push(TYPE_I32);
            }
            ("wasm:js-string", "concat") => {
                write_leb128_u32(&mut out, 2);
                out.push(TYPE_EXTERNREF);
                out.push(TYPE_EXTERNREF);
                write_leb128_u32(&mut out, 1);
                out.push(TYPE_EXTERNREF);
            }
            // wasm:js-number test: (externref) -> i32
            ("wasm:js-number", "test") => {
                write_leb128_u32(&mut out, 1);
                out.push(TYPE_EXTERNREF);
                write_leb128_u32(&mut out, 1);
                out.push(TYPE_I32);
            }
            // wasm:js-string operations with typed signatures
            ("wasm:js-string", "equals") | ("wasm:js-string", "compare") => {
                write_leb128_u32(&mut out, 2);
                out.push(TYPE_EXTERNREF);
                out.push(TYPE_EXTERNREF);
                write_leb128_u32(&mut out, 1);
                out.push(TYPE_I32);
            }
            ("wasm:js-string", "length") => {
                write_leb128_u32(&mut out, 1);
                out.push(TYPE_EXTERNREF);
                write_leb128_u32(&mut out, 1);
                out.push(TYPE_I32);
            }
            ("wasm:js-string", "charCodeAt") => {
                write_leb128_u32(&mut out, 2);
                out.push(TYPE_EXTERNREF);
                out.push(TYPE_I32);
                write_leb128_u32(&mut out, 1);
                out.push(TYPE_I32);
            }
            ("wasm:js-string", "fromCharCode") => {
                write_leb128_u32(&mut out, 1);
                out.push(TYPE_I32);
                write_leb128_u32(&mut out, 1);
                out.push(TYPE_EXTERNREF);
            }
            ("wasm:js-string", "substring") => {
                write_leb128_u32(&mut out, 3);
                out.push(TYPE_EXTERNREF);
                out.push(TYPE_I32);
                out.push(TYPE_I32);
                write_leb128_u32(&mut out, 1);
                out.push(TYPE_EXTERNREF);
            }
            // ── js-string-builtins additions ───────────────────────────
            // cast: (externref) -> stringref / externref (validates)
            ("wasm:js-string", "cast") => {
                write_leb128_u32(&mut out, 1);
                out.push(TYPE_EXTERNREF);
                write_leb128_u32(&mut out, 1);
                out.push(TYPE_EXTERNREF);
            }
            // fromCodePoint: (i32) -> externref
            ("wasm:js-string", "fromCodePoint") => {
                write_leb128_u32(&mut out, 1);
                out.push(TYPE_I32);
                write_leb128_u32(&mut out, 1);
                out.push(TYPE_EXTERNREF);
            }
            // codePointAt: (externref, i32) -> i32
            ("wasm:js-string", "codePointAt") => {
                write_leb128_u32(&mut out, 2);
                out.push(TYPE_EXTERNREF);
                out.push(TYPE_I32);
                write_leb128_u32(&mut out, 1);
                out.push(TYPE_I32);
            }
            // intoCharCodeArray: (externref, arrayref, i32) -> i32  (we simplify
            // to (externref, externref, i32) -> i32 for externref-array host)
            ("wasm:js-string", "intoCharCodeArray") => {
                write_leb128_u32(&mut out, 3);
                out.push(TYPE_EXTERNREF);
                out.push(TYPE_EXTERNREF);
                out.push(TYPE_I32);
                write_leb128_u32(&mut out, 1);
                out.push(TYPE_I32);
            }
            // fromCharCodeArray: (arrayref, i32, i32) -> externref
            ("wasm:js-string", "fromCharCodeArray") => {
                write_leb128_u32(&mut out, 3);
                out.push(TYPE_EXTERNREF);
                out.push(TYPE_I32);
                out.push(TYPE_I32);
                write_leb128_u32(&mut out, 1);
                out.push(TYPE_EXTERNREF);
            }
            // js-primitive-builtins extensions — number → string formatting
            ("wasm:js-string", "fromI32")
            | ("wasm:js-string", "fromU32") => {
                write_leb128_u32(&mut out, 1);
                out.push(TYPE_I32);
                write_leb128_u32(&mut out, 1);
                out.push(TYPE_EXTERNREF);
            }
            ("wasm:js-string", "fromI64")
            | ("wasm:js-string", "fromU64") => {
                write_leb128_u32(&mut out, 1);
                out.push(TYPE_I64);
                write_leb128_u32(&mut out, 1);
                out.push(TYPE_EXTERNREF);
            }
            ("wasm:js-string", "fromF64") => {
                write_leb128_u32(&mut out, 1);
                out.push(TYPE_F64);
                write_leb128_u32(&mut out, 1);
                out.push(TYPE_EXTERNREF);
            }

            // ── js-number unsigned variants ────────────────────────────
            ("wasm:js-number", "fromU32") => {
                write_leb128_u32(&mut out, 1);
                out.push(TYPE_I32);
                write_leb128_u32(&mut out, 1);
                out.push(TYPE_EXTERNREF);
            }
            ("wasm:js-number", "toU32") => {
                write_leb128_u32(&mut out, 1);
                out.push(TYPE_EXTERNREF);
                write_leb128_u32(&mut out, 1);
                out.push(TYPE_I32);
            }
            ("wasm:js-number", "testI32") | ("wasm:js-number", "testU32") => {
                write_leb128_u32(&mut out, 1);
                out.push(TYPE_EXTERNREF);
                write_leb128_u32(&mut out, 1);
                out.push(TYPE_I32);
            }

            // ── js-boolean.cast — (externref) -> i32 ──────────────────
            ("wasm:js-boolean", "cast") => {
                write_leb128_u32(&mut out, 1);
                out.push(TYPE_EXTERNREF);
                write_leb128_u32(&mut out, 1);
                out.push(TYPE_I32);
            }

            // ── js-symbol ──────────────────────────────────────────────
            ("wasm:js-symbol", "test") => {
                write_leb128_u32(&mut out, 1);
                out.push(TYPE_EXTERNREF);
                write_leb128_u32(&mut out, 1);
                out.push(TYPE_I32);
            }
            ("wasm:js-symbol", "equals") => {
                write_leb128_u32(&mut out, 2);
                out.push(TYPE_EXTERNREF);
                out.push(TYPE_EXTERNREF);
                write_leb128_u32(&mut out, 1);
                out.push(TYPE_I32);
            }

            // ── js-bigint ──────────────────────────────────────────────
            ("wasm:js-bigint", "test") => {
                write_leb128_u32(&mut out, 1);
                out.push(TYPE_EXTERNREF);
                write_leb128_u32(&mut out, 1);
                out.push(TYPE_I32);
            }

            _ => {
                // Default for unknown vybe:rt calls: () -> externref
                write_leb128_u32(&mut out, 0);
                write_leb128_u32(&mut out, 1);
                out.push(TYPE_EXTERNREF);
            }
        }
    }

    for (i, chunk) in chunks.iter().enumerate() {
        let type_idx = ctx.func_type_base + import_count as u32 + i as u32;
        out.push(TYPE_FUNC);
        // WASM convention: arity params (slot 0 = first arg, no reserved callee slot).
        let param_count = chunk.arity as u32;
        write_leb128_u32(&mut out, param_count);
        for _ in 0..param_count { out.push(TYPE_EXTERNREF); }
        write_leb128_u32(&mut out, 1);
        out.push(TYPE_EXTERNREF);
        // Record first type index seen for each arity (for call_ref/call_indirect dispatch)
        ctx.func_type_by_arity.entry(chunk.arity).or_insert(type_idx);
    }

    (out, ctx)
}
