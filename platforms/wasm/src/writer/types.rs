//! WASM type section encoding with Custom Descriptors.
//!
//! Each TypeEntry from the compiler produces TWO WASM GC types:
//! 1. The described struct type (object fields as externref)
//! 2. The descriptor struct type (JS prototype + vtable methods)
//!
//! This follows the Custom Descriptors proposal:
//! proposals/custom-descriptors/proposals/custom-descriptors/Overview.md

use crate::encoding::*;
use vybe_runtime::Chunk;

// Custom Descriptors binary encoding
const CD_DESCRIPTOR: u8 = 0x4D; // (descriptor $x) prefix
const CD_DESCRIBES: u8 = 0x4C; // (describes $x) prefix
const CD_SUB_FINAL: u8 = 0x4F; // sub final

/// Context for .wasm emission — maps internal types to WASM type indices.
pub struct WasmTypeContext {
    /// type_name (lowercased) → WASM type index for the described struct
    pub struct_type_indices: std::collections::HashMap<String, u32>,
    /// The module's own type index space → WASM type index. Slot `n - 1` is
    /// the 1-based immediate a `ref.test` / `struct.new` carries, so a type
    /// reference reaches the binary without passing through a name.
    pub struct_type_by_module_index: Vec<u32>,
    /// type_name → WASM type index for the descriptor struct (vtable + proto)
    pub desc_type_indices: std::collections::HashMap<String, u32>,
    /// Absolute WASM global index of the FIRST per-class descriptor singleton,
    /// or `None` when the module declares no classes.
    ///
    /// Custom Descriptors forbids `struct.new` / `struct.new_default` from
    /// allocating a type that carries a `(descriptor …)` clause — and
    /// `build_type_context` stamps one on every class — so each allocation
    /// needs a descriptor VALUE as its last operand. That value has to be one
    /// singleton per class, not one per allocation: `ref.cast_desc_eq` compares
    /// descriptors by IDENTITY, so per-allocation descriptors would make two
    /// instances of the same class fail to match each other, and each instance
    /// would reflect its own JS prototype out of descriptor field 0.
    ///
    /// The singletons are appended AFTER every module-defined global so that
    /// the index space `chunk::global_index_space` hands the compiler is
    /// untouched — class `i` (the same ordinal `desc_type_indices` is built
    /// from) owns `desc_global_base + i`.
    pub desc_global_base: Option<u32>,
    /// Per-class descriptor VTABLE plan, in class-ordinal order — one entry
    /// per descriptor global, holding the WASM function index of every method
    /// in `TypeEntry.methods` order.
    ///
    /// ⛔ Built inside the same loop that writes the descriptor struct's field
    /// count, because `struct.new` carries NO count immediate: the operand list
    /// **is** the field count. A plan assembled anywhere else can drift from
    /// `2 + te.methods.len()` by one and the module is not subtly wrong, it is
    /// invalid. The `1 +` / `2 +` split between writer and compiler is exactly
    /// that drift, caught only because nothing read a descriptor yet.
    ///
    /// `None` in a slot means the method's chunk index was out of range: that
    /// slot encodes `ref.null func` so the struct keeps its shape rather than
    /// the whole class falling back to a null vtable.
    pub desc_vtable: Vec<Vec<Option<u32>>>,
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
    /// Declared functype spelling → its type index. What `call_indirect` needs:
    /// arity cannot pick between two same-arity types.
    pub func_type_by_signature: std::collections::HashMap<String, u32>,
    /// Function indices whose DECLARED result is a bare `i32`/`f64` rather than
    /// the externref every other value carries, mapped to that valtype byte.
    ///
    /// ⛔ THE VALUE ABI IS EXTERNREF; THE SPEC SIGNATURES ARE NOT. A proposal
    /// builtin like `wasm:js-string.test` really does return i32, and once the
    /// import is typed truthfully its result can no longer be handed to an
    /// externref consumer unconverted — an `if (result externref)` arm ending
    /// in `call $length` is "expected externref, got i32". The call site has to
    /// box, so it has to know; this is how it knows.
    pub raw_result_funcs: std::collections::HashMap<u32, u8>,

    /// Declared PARAM valtypes for imports that take a raw numeric operand.
    ///
    /// ⛔ THE MIRROR OF `raw_result_funcs`, AND IT WAS MISSING. A truthful
    /// signature has two halves and so does the ABI reconciliation: results
    /// are boxed onto the externref value ABI on the way out, and operands
    /// have to be unboxed on the way in. `wasm:js-string.fromF64` is declared
    /// `(param f64)` and every caller pushed a BOXED number, because this
    /// writer boxes every f64 it produces — so the argument arrived as an
    /// externref and the module was invalid at every single call site.
    /// Only imports with at least one non-externref param appear here.
    pub raw_param_funcs: std::collections::HashMap<u32, Vec<u8>>,
    /// String-constant text → its GLOBAL index. Needed to put a property NAME
    /// on the stack for a dynamic (typeidx 0) property access, which lowers to
    /// a host call rather than a struct op. Computed here from the chunks so no
    /// call site has to thread it in.
    pub string_const_global: std::collections::HashMap<String, u32>,
    /// `externref^M -> externref^N` functype indices keyed by
    /// (param_count, result_count). Referenced by multi-value block
    /// headers as their s33 typeidx `blocktype`, AND by the
    /// call_indirect/return_call_indirect `(type $sig)` annotations
    /// (which must match the callee's functype exactly, result count
    /// included — first-seen-arity lookup is not exact).
    pub block_type_by_results: std::collections::HashMap<(u8, u8), u32>,
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
    /// Continuation tag signature/name → emitted Wasm tag index.
    pub continuation_tag_indices: std::collections::HashMap<(String, String, String), u32>,
    /// Type indices for typed continuation tags, in tag-section order.
    pub continuation_tag_type_indices: Vec<u32>,
    /// Whether any `CONT_NEW` / `SUSPEND` / `RESUME` / `SWITCH` op
    /// was observed in the bytecode. Drives whether we emit the
    /// continuation type, the suspend tag, and the tag-section entry.
    pub uses_stack_switching: bool,
}

pub(crate) fn continuation_tag_key(
    tag: &vybe_runtime::chunk::ContinuationTag,
) -> (String, String, String) {
    (
        tag.name.clone(),
        tag.yield_type.clone(),
        tag.resume_type.clone(),
    )
}

pub(crate) fn continuation_tag_valtype(type_name: &str) -> u8 {
    match type_name.to_ascii_lowercase().as_str() {
        "i32" | "bool" | "boolean" => TYPE_I32,
        "i64" => TYPE_I64,
        "f32" => TYPE_F32,
        "f64" | "number" => TYPE_F64,
        _ => TYPE_EXTERNREF,
    }
}

fn collect_continuation_tags(chunks: &[Chunk]) -> Vec<vybe_runtime::chunk::ContinuationTag> {
    let mut seen = std::collections::HashSet::new();
    let mut tags = Vec::new();
    for chunk in chunks {
        for tag in &chunk.continuation_tags {
            let key = continuation_tag_key(tag);
            if seen.insert(key) {
                tags.push(tag.clone());
            }
        }
    }
    tags
}

impl WasmTypeContext {
    /// Look up the WASM type index for a described struct type by name.
    pub fn struct_type(&self, name: &str) -> Option<u32> {
        self.struct_type_indices.get(name).copied()
    }

    /// Look up the WASM type index for a descriptor type by name.
    /// The WASM type index for a 1-based module type index — the immediate
    /// form, no name involved.
    pub fn struct_type_by_index(&self, module_index: u32) -> Option<u32> {
        if module_index == 0 {
            return None;
        }
        self.struct_type_by_module_index
            .get(module_index as usize - 1)
            .copied()
    }

    pub fn desc_type(&self, name: &str) -> Option<u32> {
        self.desc_type_indices.get(name).copied()
    }

    /// How many per-class descriptor singletons this module declares — one per
    /// class, so one per described/descriptor pair.
    pub fn descriptor_global_count(&self) -> u32 {
        self.desc_type_indices.len() as u32
    }

    /// The global holding class `name`'s descriptor singleton.
    ///
    /// Derived from the descriptor TYPE index rather than a second map, so the
    /// two cannot drift: `build_type_context` lays each class out as the pair
    /// `(2i, 2i+1)`, so `desc_type_indices[name] == 2i + 1` and the class
    /// ordinal is `(idx - 1) / 2`.
    pub fn desc_global(&self, name: &str) -> Option<u32> {
        let base = self.desc_global_base?;
        let desc_idx = self.desc_type(name)?;
        Some(base + (desc_idx - 1) / 2)
    }

    /// The descriptor singleton for a class given its DESCRIBED wasm type
    /// index — the index that actually reaches the binary.
    ///
    /// Use this when the type was resolved rather than named: the dynamic
    /// `struct.new` path guesses a wasm type from the field count and never
    /// holds a module index, but the guess can still land on a
    /// descriptor-carrying class. Described types are the even members of the
    /// `(2i, 2i+1)` layout, so an odd index (a descriptor struct, which has no
    /// descriptor of its own) or one past the pairs correctly answers `None`.
    pub fn desc_global_for_described(&self, described_idx: u32) -> Option<u32> {
        let base = self.desc_global_base?;
        if described_idx % 2 != 0 {
            return None;
        }
        let ordinal = described_idx / 2;
        if ordinal as usize >= self.struct_type_by_module_index.len() {
            return None;
        }
        Some(base + ordinal)
    }

    /// The descriptor singleton for a 1-based MODULE type index.
    ///
    /// ⚠ Prefer this over the by-name form inside the code encoder. The type
    /// table lives on `chunks[0]` ONLY, and `encode_code_section` walks every
    /// chunk, so a name lookup through `chunk.types` silently misses in every
    /// chunk but the first — which is where constructors live.
    pub fn desc_global_by_index(&self, module_index: u32) -> Option<u32> {
        let base = self.desc_global_base?;
        if module_index == 0 || module_index as usize > self.struct_type_by_module_index.len() {
            return None;
        }
        Some(base + module_index - 1)
    }

    /// Look up the field index for a field name within a struct type.
    pub fn field_index(&self, type_name: &str, field_name: &str) -> Option<u32> {
        let fields = self.struct_fields.get(type_name)?;
        fields
            .iter()
            .position(|f| f == field_name)
            .map(|i| i as u32)
    }
}

/// A declared value type's binary encoding. Anything unrecognised — a GC
/// reference spelling this table does not carry — falls back to `externref`,
/// which is what the whole section used to be, so an unknown type degrades to
/// today's behaviour rather than emitting a wrong byte.
fn val_type_byte(spelling: &str) -> u8 {
    match spelling.trim() {
        "i32" => 0x7F,
        "i64" => 0x7E,
        "f32" => 0x7D,
        "f64" => 0x7C,
        "v128" => 0x7B,
        "funcref" => 0x70,
        _ => TYPE_EXTERNREF,
    }
}

/// Build the type context and encode the type section.
/// Layout: [rec group: (described struct + descriptor struct) per TypeEntry] [array type] [function types]
/// Ask the owning proposal module for a `(module, name)` pair's exact WASM
/// signature. Returns `false` when no proposal claims the pair, leaving the
/// caller to supply its own fallback — the `TYPE_FUNC` tag byte has already
/// been pushed either way.
///
/// ⛔ BOTH import tables must ask this. A builtin is the same function whether
/// it arrives as a host import or a runtime import; typing one table by arity
/// and the other by signature makes the SAME callee two different types.
fn write_proposal_signature(out: &mut Vec<u8>, module: &str, name: &str) -> bool {
    if module == crate::writer::builtins::canon_builtins::MODULE {
        crate::writer::builtins::canon_builtins::write_signature(out, name)
    } else if module == crate::writer::builtins::js_string_builtins::MODULE {
        crate::writer::builtins::js_string_builtins::write_signature(out, name)
    } else if module == crate::writer::builtins::js_object_builtins::MODULE {
        crate::writer::builtins::js_object_builtins::write_signature(out, name)
    } else {
        crate::writer::builtins::js_primitive_builtins::write_signature(out, module, name)
    }
}

/// Decode the param valtypes of the functype just appended at `start`
/// (the byte after the `TYPE_FUNC` tag). Single-byte counts are not assumed.
fn signature_params(out: &[u8], start: usize) -> Vec<u8> {
    let mut i = start;
    let mut count = 0u32;
    let mut shift = 0u32;
    loop {
        let Some(&b) = out.get(i) else { return Vec::new() };
        i += 1;
        count |= ((b & 0x7f) as u32) << shift;
        if b & 0x80 == 0 {
            break;
        }
        shift += 7;
    }
    let end = i + count as usize;
    if end > out.len() {
        return Vec::new();
    }
    out[i..end].to_vec()
}

/// Record an import whose declared params are not all `externref`, so the
/// `call` site can unbox them. i64 params are DELIBERATELY not recorded:
/// there is no i64 unbox — our value ABI is externref and js-primitive-builtins
/// exposes only `wasm:js-bigint.test`, the i64 conversions having been removed
/// from the proposal. Recording one would produce a call site that unboxes to
/// i32 and silently truncates. `wasm:js-string.fromI64`/`fromU64` are the only
/// two, and no emitter in the tree calls either.
fn record_raw_params(ctx: &mut WasmTypeContext, func_idx: u32, out: &[u8], sig_start: usize) {
    let params = signature_params(out, sig_start);
    if params.iter().any(|&t| t == TYPE_I64) {
        return;
    }
    if params.iter().any(|&t| t != TYPE_EXTERNREF) {
        ctx.raw_param_funcs.insert(func_idx, params);
    }
}

pub fn build_type_context(
    chunks: &[Chunk],
    import_count: usize,
    rt_imports: &[(&str, &str)],
) -> (Vec<u8>, WasmTypeContext) {
    let mut out = Vec::new();
    // Global index order is `rt_globals()` first, then the string constants —
    // see `sections::encode_import_section`. Computed from the chunks so the
    // map costs no call site a new parameter.
    let string_const_base = crate::writer::sections::rt_globals().len() as u32;
    let string_const_global: std::collections::HashMap<String, u32> =
        crate::writer::sections::collect_string_constants(chunks)
            .into_iter()
            .enumerate()
            .map(|(i, text)| (text, string_const_base + i as u32))
            .collect();
    let mut ctx = WasmTypeContext {
        struct_type_indices: std::collections::HashMap::new(),
        struct_type_by_module_index: Vec::new(),
        desc_type_indices: std::collections::HashMap::new(),
        desc_global_base: None,
        desc_vtable: Vec::new(),
        struct_fields: std::collections::HashMap::new(),
        array_type_idx: 0,
        string_array_type_idx: 0,
        byte_array_type_idx: 0,
        func_type_base: 0,
        gc_type_count: 0,
        func_type_by_arity: std::collections::HashMap::new(),
        func_type_by_signature: std::collections::HashMap::new(),
        raw_result_funcs: std::collections::HashMap::new(),
        raw_param_funcs: std::collections::HashMap::new(),
        string_const_global,
        block_type_by_results: std::collections::HashMap::new(),
        exception_type_idx: 0,
        suspend_tag_type_idx: 0,
        continuation_type_idx: 0,
        continuation_tag_indices: std::collections::HashMap::new(),
        continuation_tag_type_indices: Vec::new(),
        uses_stack_switching: false,
    };

    // Collect TypeEntry definitions from chunk 0.
    //
    // ⛔ DESCRIPTOR ROWS ARE NOT CLASSES — skip them here.
    //
    // Each entry below emits a `(described 2i, descriptor 2i+1)` PAIR, so a
    // row that IS a descriptor would get a pair of its own: `#desc base`
    // acquiring `#desc base__desc`, doubling the type section and leaving the
    // real descriptor unreferenced. `describes_index` is a back-pointer (0 for
    // an ordinary class row), which is what makes the test exact rather than a
    // guess at the row's name.
    let type_entries: Vec<&vybe_runtime::chunk::TypeEntry> = chunks
        .first()
        .map(|c| c.types.iter().filter(|t| t.describes_index == 0).collect())
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

    // Pre-scan chunks for BLOCK/LOOP/IF blocktypes needing a typeidx: any
    // (param_count, result_count) pair beyond the single-byte spec forms
    // (0,0)=void and (0,1)=one valtype. Each distinct pair gets its own
    // `externref^M -> externref^N` function type, referenced as an s33
    // typeidx blocktype at emission.
    let mut block_result_counts: std::collections::BTreeSet<(u8, u8)> =
        std::collections::BTreeSet::new();
    // A tag's type is `externref^arity -> ()`, the same shape a blocktype of
    // (arity, 0) has — so declaring one here gives the tag section a typeidx to
    // point at. Without this every tag had to borrow the single one-param
    // exception type, which is why a 2-ary tag could not be expressed at all.
    for chunk in chunks {
        for tag in &chunk.tags {
            block_result_counts.insert((tag.arity, 0));
        }
    }
    for chunk in chunks {
        let code = &chunk.code;
        let mut bip = 0;
        while bip + 3 < code.len() {
            let g = ((code[bip] as u16) << 8) | code[bip + 1] as u16;
            let s = ((code[bip + 2] as u16) << 8) | code[bip + 3] as u16;
            if let Some(op) = vybe_runtime::opcode::Op::decode(g, s) {
                if op == vybe_runtime::opcode::Op::BLOCK
                    || op == vybe_runtime::opcode::Op::LOOP
                    || op == vybe_runtime::opcode::Op::IF
                {
                    // Layout: group (2) + sub (2) + params (1) + results (1).
                    if bip + 5 < code.len() {
                        let params = code[bip + 4];
                        let results = code[bip + 5];
                        if params > 0 || results >= 2 {
                            block_result_counts.insert((params, results));
                        }
                    }
                } else if op == vybe_runtime::opcode::Op::CALL_INDIRECT
                    || op == vybe_runtime::opcode::Op::RETURN_CALL_INDIRECT
                {
                    // Immediates: argc (1) + tableidx (1) + results (1).
                    // The call's `(type $sig)` annotation must match the
                    // callee's functype EXACTLY — including the result
                    // count — so register the (params, results) pair.
                    // Unlike blocktypes, (0,0)/(0,1) have no shorthand here.
                    if bip + 6 < code.len() {
                        block_result_counts.insert((code[bip + 4], code[bip + 6]));
                    }
                } else if op == vybe_runtime::opcode::Op::CALL_REF
                    || op == vybe_runtime::opcode::Op::RETURN_CALL
                    || op == vybe_runtime::opcode::Op::RETURN_CALL_REF
                {
                    // Immediates: argc (1) + results (1). Compilers emit
                    // results=1 (uniform boxed ABI); reader-ingested calls
                    // carry their functype's true result count.
                    if bip + 5 < code.len() {
                        block_result_counts.insert((code[bip + 4], code[bip + 5]));
                    }
                }
                bip += crate::writer::code::opcode_size(op, code, bip);
            } else {
                bip += 4;
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
        while bip + 3 < code.len() {
            let g = ((code[bip] as u16) << 8) | code[bip + 1] as u16;
            let s = ((code[bip + 2] as u16) << 8) | code[bip + 3] as u16;
            if let Some(op) = vybe_runtime::opcode::Op::decode(g, s) {
                if matches!(op,
                    o if o == vybe_runtime::opcode::Op::CONT_NEW
                      || o == vybe_runtime::opcode::Op::CONT_BIND
                      || o == vybe_runtime::opcode::Op::SUSPEND
                      || o == vybe_runtime::opcode::Op::RESUME
                      || o == vybe_runtime::opcode::Op::RESUME_THROW
                      || o == vybe_runtime::opcode::Op::SWITCH
                ) {
                    return true;
                }
                bip += crate::writer::code::opcode_size(op, code, bip);
            } else {
                bip += 4;
            }
        }
        false
    });
    ctx.uses_stack_switching = uses_stack_switching;
    let continuation_tags = collect_continuation_tags(chunks);
    let typed_continuation_type_count = continuation_tags.len() as u32;
    let stack_switching_type_count: u32 = if uses_stack_switching {
        2 + typed_continuation_type_count
    } else {
        typed_continuation_type_count
    };

    // Index layout: [gc] [func] [block] [suspend_tag_func] [continuation] [typed continuation tag funcs] [exception]
    let ss_base = gc_type_count + func_count + block_type_count;
    if uses_stack_switching {
        ctx.suspend_tag_type_idx = ss_base;
        ctx.continuation_type_idx = ss_base + 1;
    }
    let typed_tag_type_base = ss_base + if uses_stack_switching { 2 } else { 0 };
    for (i, tag) in continuation_tags.iter().enumerate() {
        ctx.continuation_tag_indices
            .insert(continuation_tag_key(tag), 2 + i as u32);
        ctx.continuation_tag_type_indices
            .push(typed_tag_type_base + i as u32);
    }
    ctx.exception_type_idx = ss_base + stack_switching_type_count;

    // Populate the name→typeidx maps up front so the supertype link in
    // each subtype can resolve a parent that appears later in the same
    // rec group.
    // ⚠ `struct_type_by_module_index` is keyed by MODULE index, and the module
    // index space includes the descriptor rows that `type_entries` filtered
    // out. Building it from the filtered list would shift every entry past the
    // first descriptor row, so it is built separately over ALL rows below —
    // a class row answers with its described index, a descriptor row with its
    // paired `2i+1`, resolved through the `describes_index` back-pointer.
    {
        let all_rows: &[vybe_runtime::chunk::TypeEntry] =
            chunks.first().map(|c| c.types.as_slice()).unwrap_or(&[]);
        // Module index (0-based here) → class ordinal, for class rows only.
        let mut class_ordinal: Vec<Option<u32>> = vec![None; all_rows.len()];
        let mut next = 0u32;
        for (i, t) in all_rows.iter().enumerate() {
            if t.describes_index == 0 {
                class_ordinal[i] = Some(next);
                next += 1;
            }
        }
        ctx.struct_type_by_module_index = all_rows
            .iter()
            .enumerate()
            .map(|(i, t)| {
                if t.describes_index == 0 {
                    class_ordinal[i].map_or(0, |k| k * 2)
                } else {
                    // `describes_index` is 1-BASED into the same table.
                    class_ordinal
                        .get(t.describes_index as usize - 1)
                        .copied()
                        .flatten()
                        .map_or(0, |k| k * 2 + 1)
                }
            })
            .collect();
        // A descriptor row is nameable too — `struct.new #desc <C>` has to
        // resolve to the descriptor struct, not to nothing.
        for (i, t) in all_rows.iter().enumerate() {
            if t.describes_index != 0 {
                if let Some(&idx) = ctx.struct_type_by_module_index.get(i) {
                    ctx.struct_type_indices.insert(t.name.clone(), idx);
                    ctx.struct_fields.insert(t.name.clone(), t.fields.clone());
                }
            }
        }
    }
    for (i, te) in type_entries.iter().enumerate() {
        let described_idx = (i as u32) * 2;
        let descriptor_idx = (i as u32) * 2 + 1;
        // ⛔ NOT `to_lowercase()`. These maps used to fold the key
        // unconditionally, which made `class Foo` and `class FOO` ONE entry:
        // the type section declared two described/descriptor pairs while the
        // global section emitted ONE singleton, `FOO`'s descriptor type had no
        // singleton at all, and `desc_global("FOO")` handed back `Foo`'s global
        // typed `(ref (exact 1))` — an INVALID module, and under the proposal
        // `ref.cast_desc_eq` would also test the two classes as each other.
        //
        // It is the "a QUERY folded UPSTREAM defeats it" shape from
        // `casesensitivityplan.md`, at a WRITE site, where it cannot be undone.
        // The conditional fold could never reach it because the fold happened at
        // the KEY. Nothing is lost by removing it: the compiler's `canon()`
        // already folds names for a case-insensitive language BEFORE they get
        // here, so those keys arrive uniform, while a case-sensitive language
        // keeps the distinction its source made. Found by Fathom against V8.
        ctx.struct_type_indices
            .insert(te.name.clone(), described_idx);
        ctx.desc_type_indices
            .insert(te.name.clone(), descriptor_idx);
        ctx.struct_fields.insert(te.name.clone(), te.fields.clone());
    }

    // Types with children must be left "open" (`sub`) rather than
    // `sub final` so their subtypes can extend them.
    // Keyed by the declared supertype INDEX, so two same-named-but-distinct
    // types can no longer be conflated into one "has children" answer.
    let mut has_children: std::collections::HashSet<u16> = std::collections::HashSet::new();
    for te in &type_entries {
        if te.parent_index != 0 {
            has_children.insert(te.parent_index);
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

            // Opening byte: `sub final` (0x4F) if no subtype extends this,
            // else `sub` (0x50) leaving the type open for extension.
            let described_final = !has_children.contains(&(i as u16 + 1));
            let sub_byte = if described_final { CD_SUB_FINAL } else { 0x50 };

            // Described struct subtype. Supertype count = 1 when parent
            // is named and resolvable, else 0. Parent's described typeidx
            // is the supertype link (the descriptor-struct side mirrors
            // by linking to the parent's descriptor).
            out.push(sub_byte);
            if let Some(parent_idx) = ctx.struct_type_by_index(te.parent_index as u32) {
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
            // Each entry emits a described/descriptor PAIR, so the parent's
            // descriptor sits one index past its described type.
            if let Some(parent_idx) = ctx.struct_type_by_index(te.parent_index as u32) {
                write_leb128_u32(&mut out, 1);
                write_leb128_u32(&mut out, parent_idx + 1);
            } else {
                write_leb128_u32(&mut out, 0);
            }
            out.push(CD_DESCRIBES);
            write_leb128_u32(&mut out, described_idx);
            out.push(GC_STRUCT);
            // TWO externref slots before the funcrefs, matching the compiler's
            // row (`append_descriptor_type_rows`: `__desc_proto`,
            // `__desc_props`, then methods in table order).
            //
            // ⛔ This emitted `1 + methods` and left `__desc_props` out
            // entirely, so the two halves disagreed on the layout of the same
            // struct: method `k` sat at compiler index `2+k` and writer index
            // `1+k`. Latent only because nothing reads a descriptor yet — it
            // detonates on the FIRST reader, which is what a field-0/vtable
            // writer is. Field 1's CONTENT is Cairn's; the SLOT has to exist
            // here or neither half can be written.
            let desc_field_count = 2 + te.methods.len();
            write_leb128_u32(&mut out, desc_field_count as u32);
            // field 0 — JS prototype; field 1 — property metadata. BOTH
            // MUTABLE, and the mutability is what makes them fillable at all.
            //
            // Their values are runtime objects, so they cannot appear in the
            // constant init expression that allocates the singleton, and every
            // descriptor field being `GC_IMMUT` meant there was no legal
            // instruction that could ever write them afterwards — the
            // descriptor was sealed empty. `struct.set` needs the field to be
            // mutable.
            //
            // It also sidesteps the ordering problem a `struct.new` runs into:
            // the descriptor's field COUNT depends on the merged method list,
            // which is only final after all emission, whereas `struct.set` of
            // field 0 needs no count.
            //
            // ⚠ Mutability here is OURS to choose — the proposal makes
            // descriptors ordinary structs and does not require immutable
            // fields. The identity guarantee comes from WHAT is stored (the
            // very object `C.prototype` is), not from the field being
            // read-only. Sub/super descriptor field mutability must MATCH for
            // the prefix rule, and it does: every class gets the same shape.
            for _ in 0..2 {
                out.push(TYPE_EXTERNREF);
                out.push(GC_MUT);
            }
            // The vtable stays IMMUTABLE. Fields 0/1 had to become mutable
            // because their values are runtime objects and so cannot appear in
            // a constant init expression — but a method's `ref.func` IS a
            // constant instruction, so the whole vtable can be supplied to the
            // `struct.new` that allocates the singleton and never needs to be
            // written again. Immutable is the honest declaration for a vtable
            // and it costs nothing here.
            for _ in &te.methods {
                out.push(TYPE_FUNCREF);
                out.push(GC_IMMUT);
            }

            // ⛔ The vtable PLAN is built here, in the same iteration that just
            // wrote `desc_field_count`, because `struct.new` carries no count
            // immediate — the operand list it is given IS the field count. Any
            // other assembly point can drift from this one, and the failure is
            // an invalid module rather than a subtle wrong answer.
            //
            // Function index = `import_count + chunk_index`, the same mapping
            // `encode_element_section` uses to fill the funcref table. That
            // segment is also what puts every chunk function in the module's
            // declared-reference set, so `ref.func` on any of them is valid in
            // a constant expression.
            ctx.desc_vtable.push(
                te.methods
                    .iter()
                    .map(|(_, chunk_idx)| {
                        (*chunk_idx < chunks.len())
                            .then(|| (import_count + *chunk_idx) as u32)
                    })
                    .collect(),
            );
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
    // Host imports use externref ABI types inferred by arity. Proposal
    // builtins own their exact signatures in their module-specific tables
    // below, including typed wasm:js-number and wasm:js-string helpers.

    // Import function types — per-import typed signatures
    // Host imports from chunk 0 — scan CALL_IMPORT bytecode to find actual arity
    let host_import_count = chunks.first().map(|c| c.imports.len()).unwrap_or(0);
    let mut host_arity: Vec<u8> = vec![0; host_import_count];
    for chunk in chunks {
        let mut ip = 0;
        while ip < chunk.code.len() {
            if ip + 3 >= chunk.code.len() {
                break;
            }
            let g = ((chunk.code[ip] as u16) << 8) | chunk.code[ip + 1] as u16;
            let s = ((chunk.code[ip + 2] as u16) << 8) | chunk.code[ip + 3] as u16;
            if let Some(op) = vybe_runtime::opcode::Op::decode(g, s) {
                if op == vybe_runtime::opcode::Op::CALL {
                    let import_idx = ((chunk.code[ip + 4] as u16) << 8) | chunk.code[ip + 5] as u16;
                    let argc = chunk.code[ip + 6];
                    if (import_idx as usize) < host_import_count {
                        host_arity[import_idx as usize] = host_arity[import_idx as usize].max(argc);
                    }
                }
                ip += crate::writer::code::opcode_size(op, &chunk.code, ip);
            } else {
                ip += 4;
            }
        }
    }
    for i in 0..host_import_count {
        out.push(TYPE_FUNC);
        // ⛔ ARITY IS NOT A SIGNATURE. A proposal builtin reached through the
        // HOST import table is the same function as one reached through the
        // runtime table — `wasm:js-string.test` returns i32 either way — but
        // this loop used to type every host import `(externref…) -> externref`
        // purely from a bytecode arity scan. A `test` result then fed an
        // `if (result i32)` as an externref and V8 refused the module. Ask the
        // owning proposal module for the real signature; the scan is only the
        // fallback for genuinely untyped host functions.
        let (module, name) = {
            let imp = &chunks[0].imports[i];
            (imp.module.as_str(), imp.name.as_str())
        };
        let sig_start = out.len();
        if write_proposal_signature(&mut out, module, name) {
            // Single-result signatures throughout, so the last byte written IS
            // the result valtype.
            if let Some(&vt) = out.last()
                && (vt == TYPE_I32 || vt == TYPE_F64)
            {
                ctx.raw_result_funcs.insert(i as u32, vt);
            }
            record_raw_params(&mut ctx, i as u32, &out, sig_start);
        } else {
            let argc = host_arity[i];
            write_leb128_u32(&mut out, argc as u32);
            for _ in 0..argc {
                out.push(TYPE_EXTERNREF);
            }
            write_leb128_u32(&mut out, 1);
            out.push(TYPE_EXTERNREF);
        }
    }
    // Runtime imports: each proposal module owns the signatures for
    // its own (module, name) pairs. Query them in order, falling back
    // to `(… ) -> externref` for anything unrecognised.
    for (rt_i, &(module, name)) in rt_imports.iter().enumerate() {
        let func_idx = (host_import_count + rt_i) as u32;
        out.push(TYPE_FUNC);
        let sig_start = out.len();
        if write_proposal_signature(&mut out, module, name) {
            if let Some(&vt) = out.last()
                && (vt == TYPE_I32 || vt == TYPE_F64)
            {
                ctx.raw_result_funcs.insert(func_idx, vt);
            }
            record_raw_params(&mut ctx, func_idx, &out, sig_start);
        } else {
            // ⛔ THE DEFAULT IS `() -> externref` — ZERO PARAMETERS. Anything
            // reaching it that is really CALLED with arguments would declare a
            // signature it does not have. Every module whose imports the writer
            // itself emits owns a signature table above; this is only for pairs
            // no proposal claims.
            write_leb128_u32(&mut out, 0);
            write_leb128_u32(&mut out, 1);
            out.push(TYPE_EXTERNREF);
        }
    }

    for (i, chunk) in chunks.iter().enumerate() {
        let type_idx = ctx.func_type_base + import_count as u32 + i as u32;
        out.push(TYPE_FUNC);
        // ⛔ THE DECLARED SIGNATURE WAS BEING DISCARDED HERE. This loop emitted
        // `externref` for every parameter and result, from `chunk.arity` — a
        // COUNT — while `chunk.func_sig` sat beside it holding the declared
        // value types. So `(func (param i32) (param i32) (result i32))` reached
        // the binary as `(param externref externref) (result externref)` with a
        // body that then ran `i32.add` on them: not a missing check, a module
        // that is genuinely ill-typed as emitted. Confirmed at the bytes by
        // Vesper — the declared pattern `60 02 7F 7C 01 7F` appears ZERO times
        // in a module that declares it, while the all-externref pattern appears
        // nine.
        //
        // `func_sig` is `None` for every chunk that is not a wast function, and
        // those keep the externref shape they had — a dynamic language has no
        // declared wasm signature to emit.
        let declared: Option<(Vec<String>, Vec<String>)> = chunk.func_sig.as_ref().map(|(p, r)| {
            let split = |s: &String| {
                s.split(',')
                    .map(str::trim)
                    .filter(|t| !t.is_empty())
                    .map(str::to_string)
                    .collect::<Vec<_>>()
            };
            (split(p), split(r))
        });
        let signature = match &declared {
            Some((p, r)) => format!("{}->{}", p.join(","), r.join(",")),
            None => String::new(),
        };
        match &declared {
            Some((params, results)) => {
                write_leb128_u32(&mut out, params.len() as u32);
                for t in params {
                    out.push(val_type_byte(t));
                }
                write_leb128_u32(&mut out, results.len() as u32);
                for t in results {
                    out.push(val_type_byte(t));
                }
            }
            None => {
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
            }
        }
        // Record first type index seen for each arity (for call_ref/call_indirect dispatch)
        ctx.func_type_by_arity
            .entry(chunk.arity)
            .or_insert(type_idx);
        // ⛔ AND BY SIGNATURE, which is what `call_indirect` needs. Keying only
        // by arity is `.or_insert` — first type seen per arity wins — so with
        // real functypes an argc-1 call site would be handed whichever argc-1
        // type came first, and `call_indirect`'s immediate must EQUAL the
        // callee's functype or a conforming engine traps.
        if !signature.is_empty() {
            ctx.func_type_by_signature
                .entry(signature)
                .or_insert(type_idx);
        }
    }

    // Block types: one `externref^M -> externref^N` per distinct
    // (params, results) pair found in the pre-scan. Index is recorded in
    // `ctx.block_type_by_results` for the code emitter to look up when
    // writing an s33 typeidx blocktype.
    let block_type_base = ctx.func_type_base + import_count as u32 + chunks.len() as u32;
    for (i, &(params, results)) in block_result_counts.iter().enumerate() {
        let tidx = block_type_base + i as u32;
        ctx.block_type_by_results.insert((params, results), tidx);
        out.push(TYPE_FUNC);
        write_leb128_u32(&mut out, params as u32);
        for _ in 0..params {
            out.push(TYPE_EXTERNREF);
        }
        write_leb128_u32(&mut out, results as u32);
        for _ in 0..results {
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
        out.push(crate::writer::proposals::stack_switching::CONT_TYPE_PREFIX);
        write_leb128_u32(&mut out, ctx.suspend_tag_type_idx);
    }

    for tag in &continuation_tags {
        out.push(TYPE_FUNC);
        write_leb128_u32(&mut out, 1);
        out.push(continuation_tag_valtype(&tag.yield_type));
        write_leb128_u32(&mut out, 1);
        out.push(continuation_tag_valtype(&tag.resume_type));
    }

    // Exception tag type — `(externref) -> ()` per the exception-handling
    // proposal. The tag section references this type index.
    out.push(TYPE_FUNC);
    write_leb128_u32(&mut out, 1); // 1 param
    out.push(TYPE_EXTERNREF);
    write_leb128_u32(&mut out, 0); // 0 results

    (out, ctx)
}
