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
const CD_DESCRIPTOR: u8 = 0x4D; // (descriptor $x) prefix
const CD_DESCRIBES: u8 = 0x4C; // (describes $x) prefix
const CD_SUB_FINAL: u8 = 0x4F; // sub final

/// Context for .wasm emission — maps internal types to WASM type indices.
pub struct WasmTypeContext {
    /// type_name (lowercased) → WASM type index for the described struct
    pub struct_type_indices: std::collections::HashMap<String, u32>,
    /// type_name → WASM type index for the descriptor struct (vtable + proto)
    pub desc_type_indices: std::collections::HashMap<String, u32>,
    /// type_name → vec of field names in order (for field index lookup)
    pub struct_fields: std::collections::HashMap<String, Vec<String>>,
    /// WASM type index for the dynamic array type — `(array (mut externref))`.
    /// Used for every `Value`-typed array in Vybe's uniform representation.
    pub array_type_idx: u32,
    /// WASM type index for `(array (mut i16))` — UTF-16 code-unit arrays.
    /// Strings-as-GC-array (inline string) path will reference this.
    pub string_array_type_idx: u32,
    /// WASM type index for `(array (mut i8))` — byte arrays. Backs
    /// `Uint8Array` / `Int8Array` when we inline TypedArrays.
    pub byte_array_type_idx: u32,
    /// First function type index (after GC types)
    pub func_type_base: u32,
    /// Total number of GC types (structs + descriptors + array types)
    pub gc_type_count: u32,
    /// arity → WASM type index for (externref * arity) -> externref
    pub func_type_by_arity: std::collections::HashMap<u8, u32>,
    /// Block-result count → WASM type index for `() -> externref^N`.
    /// Populated only for N >= 2; multi-value block headers reference
    /// these as their `blocktype` (signed-LEB128 typeidx).
    pub block_type_by_results: std::collections::HashMap<u8, u32>,
    /// Type index for `(externref) -> ()` — the shape required by the
    /// tag section's exception tag (exception-handling proposal).
    pub exception_type_idx: u32,
    /// Type index of the suspend/resume tag type
    /// `(tag (param externref) (result externref))` — used by the
    /// stack-switching proposal's `suspend` / `resume` when the
    /// module contains any `CONT_NEW` op. Zero if unused.
    pub suspend_tag_type_idx: u32,
    /// Continuation type index — `(cont $ft)` wrapping the shared
    /// single-arg single-result fiber function signature. Zero if
    /// the module doesn't use stack switching.
    pub continuation_type_idx: u32,
    /// Whether any `CONT_NEW` / `SUSPEND` / `RESUME` / `SWITCH` op
    /// was observed in the bytecode. Drives whether we emit the
    /// continuation type, the suspend tag, and the tag-section entry.
    pub uses_stack_switching: bool,
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
        fields
            .iter()
            .position(|f| f == &field_name.to_lowercase())
            .map(|i| i as u32)
    }
}

/// Build the type context and encode the type section.
/// Layout: [rec group: (described struct + descriptor struct) per TypeEntry] [array type] [function types]
pub fn build_type_context(
    chunks: &[Chunk],
    import_count: usize,
    rt_imports: &[(&str, &str)],
) -> (Vec<u8>, WasmTypeContext) {
    let mut out = Vec::new();
    let mut ctx = WasmTypeContext {
        struct_type_indices: std::collections::HashMap::new(),
        desc_type_indices: std::collections::HashMap::new(),
        struct_fields: std::collections::HashMap::new(),
        array_type_idx: 0,
        string_array_type_idx: 0,
        byte_array_type_idx: 0,
        func_type_base: 0,
        gc_type_count: 0,
        func_type_by_arity: std::collections::HashMap::new(),
        block_type_by_results: std::collections::HashMap::new(),
        exception_type_idx: 0,
        suspend_tag_type_idx: 0,
        continuation_type_idx: 0,
        uses_stack_switching: false,
    };

    // Collect TypeEntry definitions from chunk 0
    let type_entries: Vec<&crate::chunk::TypeEntry> = chunks
        .first()
        .map(|c| c.types.iter().collect())
        .unwrap_or_default();

    // Layout:
    // One big rec group holding every described/descriptor pair so that
    // parent/child subtype links and described↔descriptor back-pointers
    // all resolve as forward references within a single recursive block.
    // Then: 3 array types (each its own implicit singleton rec group)
    // Then: function types for imports + chunks
    // Then: N multi-value block types (one per distinct block `result_count >= 2`)
    // Then: 1 exception type `(externref) -> ()` for the tag section
    let gc_struct_pairs = type_entries.len() as u32;
    let array_count = 3u32;
    let func_count = (import_count + chunks.len()) as u32;
    let gc_type_count = gc_struct_pairs * 2 + array_count;
    ctx.gc_type_count = gc_type_count;
    ctx.array_type_idx = gc_struct_pairs * 2;
    ctx.string_array_type_idx = gc_struct_pairs * 2 + 1;
    ctx.byte_array_type_idx = gc_struct_pairs * 2 + 2;
    ctx.func_type_base = gc_type_count;

    // Pre-scan chunks for BLOCK/LOOP result counts >= 2. Each distinct
    // count gets its own `() -> externref^N` function type, which the
    // block/loop emission references as a typeidx blocktype.
    let mut block_result_counts: std::collections::BTreeSet<u8> = std::collections::BTreeSet::new();
    for chunk in chunks {
        let code = &chunk.code;
        let mut bip = 0;
        while bip + 1 < code.len() {
            if let Some(op) = crate::opcode::Op::decode(code[bip], code[bip + 1]) {
                if op == crate::opcode::Op::BLOCK
                    || op == crate::opcode::Op::LOOP
                    || op == crate::opcode::Op::IF
                {
                    // New layout: prefix (1) + sub (1) + result_count (1) = 3 bytes total.
                    if bip + 2 < code.len() {
                        let count = code[bip + 2];
                        if count >= 2 {
                            block_result_counts.insert(count);
                        }
                    }
                }
                bip += super::code::opcode_size(op, code, bip);
            } else {
                bip += 2;
            }
        }
    }
    let block_type_count = block_result_counts.len() as u32;

    // Pre-scan for stack-switching usage. Any CONT_NEW/SUSPEND/RESUME/
    // SWITCH opcode triggers the emission of:
    //   * one suspend tag type `(func (param externref) (result externref))`
    //   * one continuation type `(cont <suspend-tag-func-type>)`
    //   * the matching tag section entry
    let uses_stack_switching = chunks.iter().any(|chunk| {
        let code = &chunk.code;
        let mut bip = 0;
        while bip + 1 < code.len() {
            if let Some(op) = crate::opcode::Op::decode(code[bip], code[bip + 1]) {
                if matches!(op,
                    o if o == crate::opcode::Op::CONT_NEW
                      || o == crate::opcode::Op::CONT_NEW_TYPED
                      || o == crate::opcode::Op::CONT_BIND
                      || o == crate::opcode::Op::SUSPEND
                      || o == crate::opcode::Op::SUSPEND_TYPED
                      || o == crate::opcode::Op::RESUME
                      || o == crate::opcode::Op::RESUME_TYPED
                      || o == crate::opcode::Op::RESUME_THROW
                      || o == crate::opcode::Op::SWITCH
                ) {
                    return true;
                }
                bip += super::code::opcode_size(op, code, bip);
            } else {
                bip += 2;
            }
        }
        false
    });
    ctx.uses_stack_switching = uses_stack_switching;
    let stack_switching_type_count: u32 = if uses_stack_switching { 2 } else { 0 };

    // Index layout: [gc] [func] [block] [suspend_tag_func] [continuation] [exception]
    let ss_base = gc_type_count + func_count + block_type_count;
    if uses_stack_switching {
        ctx.suspend_tag_type_idx = ss_base;
        ctx.continuation_type_idx = ss_base + 1;
    }
    ctx.exception_type_idx = ss_base + stack_switching_type_count;

    // Populate the name→typeidx maps up front so the supertype link in
    // each subtype can resolve a parent that appears later in the same
    // rec group.
    for (i, te) in type_entries.iter().enumerate() {
        let described_idx = (i as u32) * 2;
        let descriptor_idx = (i as u32) * 2 + 1;
        let name_lower = te.name.to_lowercase();
        ctx.struct_type_indices
            .insert(name_lower.clone(), described_idx);
        ctx.desc_type_indices
            .insert(name_lower.clone(), descriptor_idx);
        ctx.struct_fields.insert(name_lower, te.fields.clone());
    }

    // Types with children must be left "open" (`sub`) rather than
    // `sub final` so their subtypes can extend them.
    let mut has_children: std::collections::HashSet<String> = std::collections::HashSet::new();
    for te in &type_entries {
        if !te.parent.is_empty() {
            has_children.insert(te.parent.to_lowercase());
        }
    }

    // Section header: number of top-level rectypes. One big rec group
    // for the struct pairs (if any) + 3 array singletons + func types +
    // 1 exception type.
    let exception_type_count = 1u32;
    // Stack-switching introduces 2 extra types when used: one func type
    // for the suspend/resume tag, one continuation type wrapping it.
    let ss_extra_types = stack_switching_type_count;
    let rec_group_count = if gc_struct_pairs > 0 { 1u32 } else { 0u32 };
    let total = rec_group_count
        + array_count
        + func_count
        + block_type_count
        + ss_extra_types
        + exception_type_count;
    write_leb128_u32(&mut out, total);

    // ── GC struct types: one rec group of (described, descriptor) pairs ──
    if gc_struct_pairs > 0 {
        out.push(GC_REC);
        write_leb128_u32(&mut out, gc_struct_pairs * 2);

        for (i, te) in type_entries.iter().enumerate() {
            let described_idx = (i as u32) * 2;
            let descriptor_idx = (i as u32) * 2 + 1;
            let name_lower = te.name.to_lowercase();

            // Opening byte: `sub final` (0x4F) if no subtype extends this,
            // else `sub` (0x50) leaving the type open for extension.
            let described_final = !has_children.contains(&name_lower);
            let sub_byte = if described_final { CD_SUB_FINAL } else { 0x50 };

            // Described struct subtype. Supertype count = 1 when parent
            // is named and resolvable, else 0. Parent's described typeidx
            // is the supertype link (the descriptor-struct side mirrors
            // by linking to the parent's descriptor).
            out.push(sub_byte);
            let parent_lower = te.parent.to_lowercase();
            if let Some(&parent_idx) = ctx.struct_type_indices.get(&parent_lower) {
                write_leb128_u32(&mut out, 1);
                write_leb128_u32(&mut out, parent_idx);
            } else {
                write_leb128_u32(&mut out, 0);
            }
            out.push(CD_DESCRIPTOR);
            write_leb128_u32(&mut out, descriptor_idx);
            out.push(GC_STRUCT);
            write_leb128_u32(&mut out, te.fields.len() as u32);
            for _ in &te.fields {
                out.push(TYPE_EXTERNREF);
                out.push(GC_MUT);
            }

            // Descriptor struct subtype — same supertype story but keyed
            // off the parent's descriptor index.
            out.push(sub_byte);
            if let Some(&parent_desc_idx) = ctx.desc_type_indices.get(&parent_lower) {
                write_leb128_u32(&mut out, 1);
                write_leb128_u32(&mut out, parent_desc_idx);
            } else {
                write_leb128_u32(&mut out, 0);
            }
            out.push(CD_DESCRIBES);
            write_leb128_u32(&mut out, described_idx);
            out.push(GC_STRUCT);
            let desc_field_count = 1 + te.methods.len();
            write_leb128_u32(&mut out, desc_field_count as u32);
            out.push(TYPE_EXTERNREF);
            out.push(GC_IMMUT);
            for _ in &te.methods {
                out.push(TYPE_FUNCREF);
                out.push(GC_IMMUT);
            }
        }
    }

    // ── Array types ──
    // Three flavours declared in the order the context recorded above:
    //   array_type_idx        → (array (mut externref)) — dynamic Value arrays
    //   string_array_type_idx → (array (mut i16))       — UTF-16 strings
    //   byte_array_type_idx   → (array (mut i8))        — byte / TypedArray backing
    // Packed types (i8 / i16) are only valid as array/struct field storage
    // types, not as top-level value types — they're emitted with the
    // `PACKED_*` byte tags per the GC proposal.
    out.push(GC_ARRAY);
    out.push(TYPE_EXTERNREF);
    out.push(GC_MUT);
    out.push(GC_ARRAY);
    out.push(PACKED_I16);
    out.push(GC_MUT);
    out.push(GC_ARRAY);
    out.push(PACKED_I8);
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
            if ip + 1 >= chunk.code.len() {
                break;
            }
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
        for _ in 0..argc {
            out.push(TYPE_EXTERNREF);
        }
        write_leb128_u32(&mut out, 1);
        out.push(TYPE_EXTERNREF);
    }
    // Runtime imports: each proposal module owns the signatures for
    // its own (module, name) pairs. Query them in order, falling back
    // to `(… ) -> externref` for anything unrecognised.
    for &(module, name) in rt_imports {
        out.push(TYPE_FUNC);
        let handled = if module == super::js_string_builtins::MODULE {
            super::js_string_builtins::write_signature(&mut out, name)
        } else {
            super::js_primitive_builtins::write_signature(&mut out, module, name)
        };
        if !handled {
            // Default for unknown calls: () -> externref
            write_leb128_u32(&mut out, 0);
            write_leb128_u32(&mut out, 1);
            out.push(TYPE_EXTERNREF);
        }
    }

    for (i, chunk) in chunks.iter().enumerate() {
        let type_idx = ctx.func_type_base + import_count as u32 + i as u32;
        out.push(TYPE_FUNC);
        // WASM convention: arity params (slot 0 = first arg, no reserved callee slot).
        let param_count = chunk.arity as u32;
        write_leb128_u32(&mut out, param_count);
        for _ in 0..param_count {
            out.push(TYPE_EXTERNREF);
        }
        // Multi-value proposal: chunks may return more than one externref.
        let result_count = (chunk.result_arity as u32).max(1);
        write_leb128_u32(&mut out, result_count);
        for _ in 0..result_count {
            out.push(TYPE_EXTERNREF);
        }
        // Record first type index seen for each arity (for call_ref/call_indirect dispatch)
        ctx.func_type_by_arity
            .entry(chunk.arity)
            .or_insert(type_idx);
    }

    // Block multi-value types: one `() -> externref^N` per distinct
    // block/loop `result_count >= 2` found in the pre-scan. Index is
    // recorded in `ctx.block_type_by_results` for the code emitter to
    // look up when writing a typeidx blocktype.
    let block_type_base = ctx.func_type_base + import_count as u32 + chunks.len() as u32;
    for (i, &count) in block_result_counts.iter().enumerate() {
        let tidx = block_type_base + i as u32;
        ctx.block_type_by_results.insert(count, tidx);
        out.push(TYPE_FUNC);
        write_leb128_u32(&mut out, 0); // 0 params
        write_leb128_u32(&mut out, count as u32); // N results
        for _ in 0..count {
            out.push(TYPE_EXTERNREF);
        }
    }

    // Stack-switching: suspend/resume tag func type + continuation type.
    // A suspend yields an externref and resumes with an externref, so the
    // tag's signature is `(func (param externref) (result externref))`.
    // The continuation type `(cont $ft)` wraps that func type.
    if uses_stack_switching {
        // (func (param externref) (result externref))
        out.push(TYPE_FUNC);
        write_leb128_u32(&mut out, 1); // 1 param
        out.push(TYPE_EXTERNREF);
        write_leb128_u32(&mut out, 1); // 1 result
        out.push(TYPE_EXTERNREF);
        // (cont $ft) — prefix + funcidx
        out.push(super::stack_switching::CONT_TYPE_PREFIX);
        write_leb128_u32(&mut out, ctx.suspend_tag_type_idx);
    }

    // Exception tag type — `(externref) -> ()` per the exception-handling
    // proposal. The tag section references this type index.
    out.push(TYPE_FUNC);
    write_leb128_u32(&mut out, 1); // 1 param
    out.push(TYPE_EXTERNREF);
    write_leb128_u32(&mut out, 0); // 0 results

    (out, ctx)
}
