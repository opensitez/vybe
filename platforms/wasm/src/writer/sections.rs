//! WASM module section encoding (import, function, memory, export).

use crate::encoding::*;
use vybe_runtime::Chunk;
use vybe_runtime::chunk::Import;

#[derive(Clone)]
pub struct HostImport {
    pub import: Import,
    pub arity: Option<u8>,
}

/// Collect all runtime imports needed by the chunks.
/// Returns (module, name) pairs for both vybe:rt and wasm:js-* builtins.
pub fn collect_rt_imports(chunks: &[Chunk]) -> Vec<(&'static str, &'static str)> {
    let mut needed: Vec<(&str, &str)> = Vec::new();
    let mut seen = std::collections::HashSet::new();
    if chunks.iter().any(|c| c.local_names.iter().any(|n|n.origin==vybe_runtime::chunk::LocalOrigin::Compiler && n.name.starts_with("<native-growable-array:"))) {
        needed.push(("ecma:array", "newWithLength"));
        seen.insert(("ecma:array", "newWithLength"));
    }

    // Aggregate function imports from every proposal module. Each
    // proposal owns its slice of imports and types (see the
    // per-proposal modules directly under `wasm/`).
    for name in crate::writer::builtins::js_string_builtins::IMPORTS {
        let key = (crate::writer::builtins::js_string_builtins::MODULE, *name);
        if seen.insert(key) {
            needed.push(key);
        }
    }
    for &(module, name) in crate::writer::builtins::js_primitive_builtins::FUNC_IMPORTS {
        let key = (module, name);
        if seen.insert(key) {
            needed.push(key);
        }
    }

    // ⛔ DYNAMIC PROPERTY ACCESS CANNOT BE A STRUCT OP. `STRUCT_GET`/`STRUCT_SET`
    // carry `(typeidx, name-constant)`, and typeidx 0 means "no static type" —
    // a property read by NAME on an untyped object, which is 897 of the 902
    // struct ops a plain JS class emits. The writer used to lower those to a
    // typed `struct.get` against type 0 field 0, because its name search
    // returns the `(0, 0)` sentinel on a miss. Type 0 is a REAL type, so the
    // instruction validates as well-formed and V8 rejects it for what it is:
    // `invalid field index: 0`. That fails on a two-function wast module with
    // no classes at all, so it is not a class-census problem.
    //
    // These two are what an untyped access actually is. They are host imports
    // rather than inline wasm because core wasm has no by-name property access;
    // that is the whole reason the GC proposal has typed fields and this path
    // does not use them.
    // `new` joins them for the same reason: `struct.new` with typeidx 0 is not
    // a typed allocation at all, it is "make a fresh dynamic object".
    for key in [
        ("ecma:object", "new"),
        ("ecma:object", "get"),
        ("ecma:object", "set"),
        ("ecma:array", "getValue"),
        ("ecma:array", "setValue"),
        ("ecma:array", "length"),
        ("ecma:object", "is"),
        ("ecma:function", "apply"),
    ] {
        if seen.insert(key) {
            needed.push(key);
        }
    }
    needed
}

/// Collect host function imports from every chunk into the single function
/// index space a standard Wasm module actually has.
pub fn collect_host_imports(chunks: &[Chunk]) -> Vec<HostImport> {
    let mut imports: Vec<Import> = Vec::new();
    let mut seen = std::collections::HashSet::new();
    for chunk in chunks {
        for import in &chunk.imports {
            let key = (import.module.clone(), import.name.clone());
            if seen.insert(key) {
                imports.push(import.clone());
            }
        }
    }
    let mut arities: std::collections::HashMap<(String, String), std::collections::BTreeSet<u8>> =
        std::collections::HashMap::new();
    for chunk in chunks {
        let mut ip = 0;
        while ip + 3 < chunk.code.len() {
            let g = u16::from_be_bytes([chunk.code[ip], chunk.code[ip + 1]]);
            let s = u16::from_be_bytes([chunk.code[ip + 2], chunk.code[ip + 3]]);
            if let Some(op) = vybe_runtime::Op::decode(g, s) {
                if op == vybe_runtime::Op::CALL && ip + 6 < chunk.code.len() {
                    let idx = u16::from_be_bytes([chunk.code[ip + 4], chunk.code[ip + 5]]);
                    if let Some(import) = chunk
                        .imports
                        .get(idx as usize)
                        .or_else(|| imports.get(idx as usize))
                    {
                        arities
                            .entry((import.module.clone(), import.name.clone()))
                            .or_default()
                            .insert(chunk.code[ip + 6]);
                    }
                }
                ip += crate::writer::code::opcode_size(op, &chunk.code, ip);
            } else {
                ip += 4;
            }
        }
    }
    let mut result = Vec::new();
    let mut extra = Vec::new();
    for import in imports {
        let key = (import.module.clone(), import.name.clone());
        if crate::writer::types::has_proposal_signature(&import.module, &import.name) {
            result.push(HostImport {
                import,
                arity: None,
            });
            continue;
        }
        let calls = arities.remove(&key).unwrap_or_default();
        let primary = calls.iter().next_back().copied().unwrap_or(0);
        for &arity in &calls {
            if arity != primary {
                extra.push(HostImport {
                    import: import.clone(),
                    arity: Some(arity),
                });
            }
        }
        result.push(HostImport {
            import,
            arity: Some(primary),
        });
    }
    result.extend(extra);
    result
}

pub fn host_import_index_map(
    host_imports: &[HostImport],
) -> std::collections::HashMap<(String, String, Option<u8>), usize> {
    host_imports
        .iter()
        .enumerate()
        .map(|(i, entry)| {
            (
                (
                    entry.import.module.clone(),
                    entry.import.name.clone(),
                    entry.arity,
                ),
                i,
            )
        })
        .collect()
}

/// `wasm:js-*` globals — imported as externref to give the emitter direct
/// access to the JS host's `undefined`, `true`, and `false` singletons
/// (per the js-primitive-builtins proposal — creation via global is
/// significantly cheaper than a function call per use).
/// Indices here are the WASM global-index space (separate from function
/// indices) and must match the order globals are emitted in the import
/// section.
pub const JS_GLOBAL_UNDEFINED: u32 = 0;
pub const JS_GLOBAL_TRUE: u32 = 1;
pub const JS_GLOBAL_FALSE: u32 = 2;

pub fn rt_globals() -> &'static [(&'static str, &'static str)] {
    // Sourced from the js-primitive-builtins proposal module so the
    // two places stay in lock-step. The indices here (0, 1, 2) must
    // match `JS_GLOBAL_UNDEFINED`/`TRUE`/`FALSE` constants above.
    crate::writer::builtins::js_primitive_builtins::GLOBAL_IMPORTS
}

pub fn encode_import_section(
    chunks: &[Chunk],
    host_imports: &[HostImport],
    rt_imports: &[(&str, &str)],
    func_type_base: u32,
    string_constants: &[String],
    host_globals: &[String],
) -> Vec<u8> {
    let mut out = Vec::new();
    let globals = rt_globals();
    let total = host_imports.len()
        + rt_imports.len()
        + globals.len()
        + string_constants.len()
        + host_globals.len();
    write_leb128_u32(&mut out, total as u32);

    // Host imports in module-level function-index order.
    for (i, entry) in host_imports.iter().enumerate() {
        let import = &entry.import;
        write_name(&mut out, &import.module);
        write_name(&mut out, &import.name);
        out.push(0x00); // func import
        write_leb128_u32(&mut out, func_type_base + i as u32);
    }

    // Runtime + builtin function imports (mixed modules: vybe:rt, wasm:js-*, …)
    for (i, (module, name)) in rt_imports.iter().enumerate() {
        write_name(&mut out, module);
        write_name(&mut out, name);
        out.push(0x00); // func import
        write_leb128_u32(&mut out, func_type_base + (host_imports.len() + i) as u32);
    }

    // `wasm:js-*` global imports — externref, immutable. Indices follow
    // the order in `rt_globals()` (see JS_GLOBAL_UNDEFINED/TRUE/FALSE).
    for (module, name) in globals {
        write_name(&mut out, module);
        write_name(&mut out, name);
        out.push(0x03); // global import
        out.push(TYPE_EXTERNREF);
        out.push(0x00); // immutable (mut = 0)
    }

    // js-string-builtins § String constants — one immutable externref global
    // per constant, its VALUE being the import's field name. Order matches
    // `collect_string_constants`, which is also what the `global.get`
    // immediates index.
    for text in string_constants {
        write_name(&mut out, vybe_runtime::chunk::STRING_CONSTANTS_MODULE);
        write_name(&mut out, text);
        out.push(0x03); // global import
        out.push(TYPE_EXTERNREF);
        out.push(0x00); // immutable (mut = 0)
    }

    // Free globals — the bindings this module reads but never writes. Order
    // matches `collect_host_globals`, which is also what `global_map` indexes.
    for name in host_globals {
        write_name(&mut out, vybe_runtime::chunk::HOST_GLOBALS_MODULE);
        write_name(&mut out, name);
        out.push(0x03); // global import
        out.push(TYPE_EXTERNREF);
        out.push(u8::from(
            chunks
                .iter()
                .any(|c| c.global_inits.iter().any(|i| i.name == *name)),
        )); // true initializer write
    }

    out
}

pub fn encode_func_section(chunks: &[Chunk], import_count: usize, func_type_base: u32) -> Vec<u8> {
    let mut out = Vec::new();
    write_leb128_u32(&mut out, chunks.len() as u32);
    for (i, _) in chunks.iter().enumerate() {
        write_leb128_u32(&mut out, func_type_base + import_count as u32 + i as u32);
    }
    out
}

pub(super) fn encode_func_section_with_native_resume(
    chunks: &[Chunk],
    import_count: usize,
    func_type_base: u32,
    resume_type: Option<u32>,
) -> Vec<u8> {
    let mut out = Vec::new();
    write_leb128_u32(
        &mut out,
        chunks.len() as u32 + u32::from(resume_type.is_some()),
    );
    for (i, _) in chunks.iter().enumerate() {
        write_leb128_u32(&mut out, func_type_base + import_count as u32 + i as u32);
    }
    if let Some(t) = resume_type {
        write_leb128_u32(&mut out, t);
    }
    out
}

pub fn encode_memory_section() -> Vec<u8> {
    // Default: single non-shared memory, 1 page min, no max.
    encode_memory_section_with(1, None, false)
}

/// Serialize the module's declared memories in their core Wasm index order.
/// The first chunk owns the memory index space, as it does in the VM.
pub fn encode_declared_memory_section(chunks: &[Chunk]) -> Vec<u8> {
    let declared = chunks
        .first()
        .filter(|chunk| !chunk.memory_min_pages.is_empty() || !chunk.memory_is_64.is_empty());
    let Some(chunk) = declared else {
        return encode_memory_section();
    };

    let count = chunk.memory_min_pages.len().max(chunk.memory_is_64.len());
    let mut out = Vec::new();
    write_leb128_u32(&mut out, count as u32);
    for index in 0..count {
        let memory64 = chunk.memory_is_64.get(index).copied().unwrap_or(false);
        let min = chunk
            .memory_min_pages
            .get(index)
            .copied()
            .unwrap_or(u64::from(index == 0));
        let max = chunk.memory_max_pages.get(index).copied().flatten();
        out.push(u8::from(max.is_some()) | if memory64 { 0x04 } else { 0 });
        if memory64 {
            write_leb128_u64(&mut out, min);
            if let Some(max) = max {
                write_leb128_u64(&mut out, max);
            }
        } else {
            write_leb128_u32(
                &mut out,
                u32::try_from(min).expect("memory32 minimum exceeds u32"),
            );
            if let Some(max) = max {
                write_leb128_u32(
                    &mut out,
                    u32::try_from(max).expect("memory32 maximum exceeds u32"),
                );
            }
        }
    }
    out
}

pub fn module_uses_memory64(chunks: &[Chunk]) -> bool {
    // Memory64-ness is a property of the memory's declared index type
    // (`memory_is_64`), not of any opcode — memory64 adds no new opcodes.
    chunks
        .first()
        .is_some_and(|chunk| chunk.memory_is_64.iter().any(|&is64| is64))
}

/// Memory-section encoder with explicit limits and `shared` flag.
/// `shared = true` requires a max page count (enforced here) and sets
/// the limits flag byte to 0x03 per the threads proposal.
pub fn encode_memory_section_with(min_pages: u32, max_pages: Option<u32>, shared: bool) -> Vec<u8> {
    let mut out = Vec::new();
    write_leb128_u32(&mut out, 1); // one memory
    // Limits flags: bit 0 = has max, bit 1 = shared.
    let has_max = max_pages.is_some() || shared;
    let flags: u8 = (if has_max { 0x01 } else { 0x00 }) | (if shared { 0x02 } else { 0x00 });
    out.push(flags);
    write_leb128_u32(&mut out, min_pages);
    if has_max {
        let max = max_pages.unwrap_or(min_pages.max(1));
        write_leb128_u32(&mut out, max);
    }
    out
}

pub fn encode_memory64_section_with(
    min_pages: u64,
    max_pages: Option<u64>,
    shared: bool,
) -> Vec<u8> {
    let mut out = Vec::new();
    write_leb128_u32(&mut out, 1); // one memory
    let has_max = max_pages.is_some() || shared;
    let flags: u8 = 0x04 | (if has_max { 0x01 } else { 0x00 }) | (if shared { 0x02 } else { 0x00 });
    out.push(flags);
    write_leb128_u64(&mut out, min_pages);
    if has_max {
        let max = max_pages.unwrap_or(min_pages.max(1));
        write_leb128_u64(&mut out, max);
    }
    out
}

pub(crate) fn start_chunk_index(chunks: &[Chunk]) -> usize {
    chunks
        .iter()
        .position(|chunk| chunk.name == "<script>" && !chunk.code.is_empty())
        .or_else(|| {
            chunks
                .iter()
                .position(|chunk| chunk.name == "main" || chunk.name == "__main")
        })
        .unwrap_or(0)
}

pub fn encode_export_section(chunks: &[Chunk], import_count: usize) -> Vec<u8> {
    encode_export_section_with_entry(chunks, import_count, None)
}

pub(crate) fn encode_export_section_with_entry(
    chunks: &[Chunk], import_count: usize, entry_function: Option<u32>,
) -> Vec<u8> {
    let mut out = Vec::new();
    write_leb128_u32(&mut out, 2); // export memory + main func
    // Memory
    write_name(&mut out, "memory");
    out.push(0x02); // memory export
    write_leb128_u32(&mut out, 0);
    // Main function. Chunk 0 is often a script wrapper, but C-style programs
    // can leave it empty and put the real entry in `main`.
    write_name(&mut out, "_start");
    out.push(0x00); // func export
    write_leb128_u32(
        &mut out,
        entry_function.unwrap_or(import_count as u32 + start_chunk_index(chunks) as u32),
    );
    out
}

/// Every string constant the module imports, in declaration order, deduped.
///
/// js-string-builtins § String constants: each becomes
/// `(global (import "wasm:string-constants" "<text>") (ref extern))` and the
/// import's field name IS the value. Collected over **all** chunks — each
/// declares its own imports — and this exact order is what both the import
/// section and the `global.get` immediates index into. If the two ever
/// disagreed the module would still validate and read the *wrong string*.
pub fn collect_string_constants(chunks: &[Chunk]) -> Vec<String> {
    // ⛔ DO NOT ADD ENTRIES HERE. Tried and reverted 2026-08-28: appending the
    // dynamic-property names (so a typeidx-0 access could name its key) shifts
    // the GLOBAL INDEX SPACE, because imports precede defined globals and the
    // writer uses the bytecode's `GLOBAL_GET`/`GLOBAL_SET` immediate directly
    // as the wasm global index. The compiler numbered those immediates over
    // `global_imports ++ defined` WITHOUT the added names, so every defined
    // global moved and V8 answered `immutable global #12 cannot be assigned` —
    // a `global.set` landing on an imported string constant.
    //
    // The names have to be declared as real `chunk.global_imports` by the
    // COMPILER, which is the only side that can keep both numberings in step.
    let mut names = collect_global_imports(chunks, vybe_runtime::chunk::STRING_CONSTANTS_MODULE);
    if !super::generator_construct::entries(chunks).is_empty() {
        for name in ["next", "iterator", "asyncIterator", "__bound_args"] {
            if !names.iter().any(|n| n == name) {
                names.push(name.into());
            }
        }
    }
    names
}


/// The module's FREE globals, in declaration order — names it reads but never
/// writes. They are imports of the embedder, so they belong in the import
/// section; without this they were emitted as DEFINED globals initialised to
/// `ref.null extern`, which is a module that declares it owns `globalThis`.
pub fn collect_original_host_globals(chunks: &[Chunk]) -> Vec<String> {
    collect_global_imports(chunks, vybe_runtime::chunk::HOST_GLOBALS_MODULE)
}
pub fn collect_host_globals(chunks: &[Chunk]) -> Vec<String> {
    // Initializers write module-owned storage; they do not grant an import
    // contract. Explicit imports retain their identity even when initialized.
    collect_original_host_globals(chunks)
}

fn collect_global_imports(chunks: &[Chunk], module: &str) -> Vec<String> {
    let mut out: Vec<String> = Vec::new();
    for chunk in chunks {
        for import in &chunk.global_imports {
            if import.module == module && !out.iter().any(|existing| existing == &import.name) {
                out.push(import.name.clone());
            }
        }
    }
    out
}

/// Collect the module's globals.
///
/// Returns the **defined** globals (what the global section declares) and a
/// name → **absolute WASM global index** map covering both kinds. WASM numbers
/// imported globals first, so the low indices are the `rt_globals()` js-primitive
/// singletons followed by one per string constant; defined globals follow.
/// Imported globals are in the map under their `(module, name)` key, so
/// `global.get` on a string constant and on a user variable resolve through one
/// lookup — and a user global can never be assigned an imported global's index.
/// ⚠ The module's DEFINED globals, for the global section only.
///
/// This used to also return a `name -> absolute index` map that the code
/// encoder resolved `GLOBAL_GET` through. That map is gone: the compiler now
/// assigns the index itself (`globals::normalize_global_table`) over the SAME
/// ordering, defined once in `vybe_runtime::chunk::global_index_space`. Two
/// implementations that merely AGREED is exactly the failure this area already
/// had — a chunk-0 coincidence held together the whole numbering.
pub fn collect_globals(
    chunks: &[Chunk],
    string_constants: &[String],
    host_globals: &[String],
) -> Vec<String> {
    // CONSUMER of the compiler's index space, not a second construction of it.
    //
    // `chunks[0].globals` is the whole space in WASM's order —
    // `chunk::global_index_space`: the primitive imports, then one imported
    // global per string constant, then host globals, then the module's own.
    // The global SECTION declares only that last part, so this returns the
    // tail. Rebuilding it here by walking the bytecode is what this used to do,
    // and it cannot work any more: a `GLOBAL_GET` operand is a globalidx now,
    // not a constant index naming the global.
    // Native constructor protocol constants do not belong to the compiler's
    // original index space. Slice the original defined tail before native imports
    // are appended, then map all surviving original names into the native space.
    let original_strings = collect_global_imports(chunks, vybe_runtime::chunk::STRING_CONSTANTS_MODULE);
    let defined_base = vybe_runtime::chunk::global_index_space(
        &original_strings,
        &collect_original_host_globals(chunks),
        &[],
    )
    .len();
    chunks
        .first()
        .map(|c| {
            c.globals
                .iter()
                .skip(defined_base)
                .filter(|name| !host_globals.contains(name))
                .cloned()
                .collect()
        })
        .unwrap_or_default()
}

/// Encode the global section — all globals are (mut externref), initialized to null.
pub fn encode_global_section(globals: &[String]) -> Vec<u8> {
    encode_global_section_with_descriptors(globals, None)
}

/// The global section, plus one descriptor SINGLETON per class.
///
/// Custom Descriptors forbids `struct.new` / `struct.new_default` from
/// allocating a type carrying a `(descriptor …)` clause, and `types.rs` stamps
/// one on every class — so every class allocation needs a descriptor value as
/// its last operand, and that value must be one singleton per class rather than
/// one per allocation (`ref.cast_desc_eq` is identity, and descriptor field 0
/// is the JS prototype every instance of the class must share).
///
/// Each singleton is emitted as
///
/// ```text
/// (global (ref (exact $C.desc))
///   (struct.new $C.desc
///     (ref.null extern)          ;; field 0 — __desc_proto, filled by desc.set_proto
///     (ref.null extern)          ;; field 1 — __desc_props
///     (ref.func $C.m0) …))       ;; the vtable, one funcref per method
/// ```
///
/// * **Non-nullable and immutable.** `struct.new_default_desc` traps on a null
///   descriptor; typing the global `(ref …)` rather than `(ref null …)` makes
///   that trap unreachable by construction instead of by argument, and nothing
///   emits `global.set` against these (the only consumer is the `global.get` in
///   `code.rs`'s allocation path), so immutable is the true declaration.
///
///   ⛔ This briefly became mutable, on the theory that the class-definition
///   path would OVERWRITE the singleton with a `struct.new` carrying the
///   prototype. That is not what landed — field 0 is filled by `struct.set`,
///   which needs the FIELD mutable, not the global — and a mutable global whose
///   justification never shipped is worse than no comment.
///
/// * **`exact`.** The proposal types the descriptor operand
///   `(ref null (exact y))` — a non-exact reference does not satisfy it, so
///   `0x62` is load-bearing here, not decoration.
///
/// * **`struct.new` is legal on the DESCRIPTOR type, and legal HERE.** The
///   allocation prohibition is on types carrying a `descriptor` clause; a
///   descriptor struct carries `describes` and has none of its own. It is also
///   a *constant* instruction, as is `ref.func`, which is what lets the whole
///   vtable be supplied in a global initialiser rather than start-function
///   code. The functions are in the module's declared-reference set already —
///   `encode_element_section` names every chunk function in an active segment.
///
/// * **Fields 0 and 1 are left null here** and written afterwards. Their values
///   are runtime objects, so they cannot appear in a constant expression at
///   all; that is the entire reason they are the two mutable fields.
///
/// Appended AFTER every module-defined global, so the index space the compiler
/// assigned (`chunk::global_index_space`) is untouched.
pub fn encode_global_section_with_descriptors(
    globals: &[String],
    descriptors: Option<&crate::writer::types::WasmTypeContext>,
) -> Vec<u8> {
    let mut out = Vec::new();
    let desc_count = descriptors.map_or(0, |c| c.descriptor_global_count());
    write_leb128_u32(&mut out, globals.len() as u32 + desc_count);
    for _ in globals {
        out.push(0x6F); // externref
        out.push(0x01); // mutable
        // Init expr: ref.null extern
        out.push(0xD0);
        out.push(0x6F); // ref.null extern
        out.push(0x0B); // end
    }
    if desc_count > 0 {
        // In class-ordinal order — `desc_global()` derives the index from the
        // same `(2i, 2i+1)` layout, so this loop and that lookup cannot drift.
        for i in 0..desc_count {
            let desc_type_idx = i * 2 + 1;
            out.push(0x64); // (ref ht) — NON-nullable
            out.push(0x62); // (exact x)
            write_leb128_u32(&mut out, desc_type_idx);
            out.push(0x00); // immutable

            // Count and initialization both use the structural class-ordinal
            // plan. Diagnostic-name aliases must not collapse descriptor
            // singletons or shift a later class's global index.
            let vtable = descriptors.and_then(|c| c.desc_vtable.get(i as usize));
            match vtable {
                Some(methods) => {
                    // Fields 0 and 1: `ref.null extern`. Runtime objects can
                    // never be constant-expressed; `desc.set_proto` fills
                    // field 0 once the prototype exists.
                    for _ in 0..2 {
                        out.push(0xD0);
                        out.push(0x6F); // ref.null extern
                    }
                    for m in methods {
                        match m {
                            Some(func_idx) => {
                                out.push(0xD2); // ref.func
                                write_leb128_u32(&mut out, *func_idx);
                            }
                            // Unresolvable chunk index — null THIS slot only.
                            // The operand count must still match the field
                            // count exactly.
                            None => {
                                out.push(0xD0);
                                out.push(0x70); // ref.null func
                            }
                        }
                    }
                    out.push(0xFB);
                    write_leb128_u32(&mut out, 0x00); // struct.new
                    write_leb128_u32(&mut out, desc_type_idx);
                }
                None => {
                    out.push(0xFB);
                    write_leb128_u32(&mut out, 0x01); // struct.new_default
                    write_leb128_u32(&mut out, desc_type_idx);
                }
            }
            out.push(0x0B); // end
        }
    }
    out
}

pub fn encode_table_section(chunks: &[Chunk], _import_count: usize) -> Vec<u8> {
    // Single funcref table for `call_indirect` dispatch. reference-types
    // lets us declare multiple tables, but emitting one nothing uses is
    // the same kind of dead-weight noise we just pruned from imports.
    // An opt-in helper (`encode_table_section_with`) is available when a
    // chunk actually wants an extra externref table.
    let mut out = Vec::new();
    let table_size = chunks.len() as u32;
    write_leb128_u32(&mut out, 1); // 1 table
    out.push(0x70); // funcref
    out.push(0x00); // no max
    write_leb128_u32(&mut out, table_size);
    out
}

pub fn encode_table_section_with_context(
    chunks: &[Chunk], ctx: &super::types::WasmTypeContext,
) -> Vec<u8> {
    use vybe_runtime::opcode::heaptype::HeapType;
    let mut out = Vec::new();
    let script = chunks.first();
    let count = script.map_or(0, |chunk| chunk.table_min_sizes.len());
    write_leb128_u32(&mut out, count as u32 + 1);
    if let Some(script) = script {
        for (index, &min) in script.table_min_sizes.iter().enumerate() {
            let (heap, nullable, exact, shared) = script.table_element_types.get(index).copied()
                .unwrap_or((HeapType::Abstract(0x70), true, false, false));
            encode_segment_reference_type(&mut out, (heap, nullable, exact, shared), ctx);
            let max = script.table_max_sizes.get(index).copied().flatten();
            let is_64 = script.table_is_64.get(index).copied().unwrap_or(false);
            let table_shared = script.table_is_shared.get(index).copied().unwrap_or(false);
            out.push(u8::from(max.is_some()) | (u8::from(table_shared) << 1) | (u8::from(is_64) << 2));
            if is_64 {
                write_leb128_u64(&mut out, min);
                if let Some(max) = max { write_leb128_u64(&mut out, max); }
            } else {
                write_leb128_u32(&mut out, min as u32);
                if let Some(max) = max { write_leb128_u32(&mut out, max as u32); }
            }
        }
    }
    out.extend([0x70, 0]);
    write_leb128_u32(&mut out, chunks.len() as u32);
    out
}

/// Emit a table section that additionally declares `extra_externref`
/// externref-typed tables (each with zero min-size, no max). Kept
/// private to the reference-types module in documentation but exported
/// so downstream tools can opt in.
pub fn encode_table_section_with(chunks: &[Chunk], extra_externref: u32) -> Vec<u8> {
    let mut out = Vec::new();
    let table_size = chunks.len() as u32;
    let total = 1u32 + extra_externref;
    write_leb128_u32(&mut out, total);
    out.push(0x70);
    out.push(0x00);
    write_leb128_u32(&mut out, table_size);
    for _ in 0..extra_externref {
        out.push(0x6F);
        out.push(0x00);
        write_leb128_u32(&mut out, 0);
    }
    out
}

pub fn encode_table64_section_with(
    min_entries: u64,
    max_entries: Option<u64>,
    reftype: u8,
) -> Vec<u8> {
    let mut out = Vec::new();
    write_leb128_u32(&mut out, 1); // one table
    out.push(reftype);
    if let Some(max) = max_entries {
        out.push(0x05); // memory64/table64 limits with max
        write_leb128_u64(&mut out, min_entries);
        write_leb128_u64(&mut out, max);
    } else {
        out.push(0x04); // memory64/table64 min-only limits
        write_leb128_u64(&mut out, min_entries);
    }
    out
}

pub fn encode_element_section(chunks: &[Chunk], import_count: usize) -> Vec<u8> {
    let mut out = Vec::new();
    write_leb128_u32(&mut out, 1); // 1 element segment
    // Active segment for table 0, offset 0
    out.push(0x00); // flags: active, table 0, funcref
    out.push(0x41);
    write_leb128_i32(&mut out, 0); // i32.const 0 (offset)
    out.push(0x0B); // end init expr
    // Function indices
    write_leb128_u32(&mut out, chunks.len() as u32);
    for i in 0..chunks.len() {
        write_leb128_u32(&mut out, (import_count + i) as u32);
    }
    out
}

pub fn collect_data_segments(chunks: &[Chunk]) -> Vec<Vec<u8>> {
    collect_data_segments_borrowed(chunks).to_vec()
}

pub(crate) fn collect_data_segments_borrowed(chunks: &[Chunk]) -> &[Vec<u8>] {
    chunks
        .first()
        .map(|chunk| chunk.data_segments.as_slice())
        .unwrap_or(&[])
}

pub fn encode_data_count_section(data_segments: &[Vec<u8>]) -> Vec<u8> {
    let mut out = Vec::new();
    write_leb128_u32(&mut out, data_segments.len() as u32);
    out
}

pub fn encode_data_section(data_segments: &[Vec<u8>]) -> Vec<u8> {
    let mut out = Vec::new();
    write_leb128_u32(&mut out, data_segments.len() as u32);
    for bytes in data_segments {
        out.push(0x01); // passive data segment; code uses memory.init/data.drop.
        write_leb128_u32(&mut out, bytes.len() as u32);
        out.extend_from_slice(bytes);
    }
    out
}

/// Preserve segment indices and instantiation mode. Legacy/source payloads
/// without active declarations retain their existing passive representation.
pub(crate) fn encode_data_section_with_initializers(
    chunks: &[Chunk],
    data_segments: &[Vec<u8>],
    ctx: &super::types::WasmTypeContext,
) -> Vec<u8> {
    if let Some(module) = chunks.first() {
        if let Some(segments) = &module.native_segments {
            use vybe_runtime::chunk::NativeSegmentMode;
            assert_eq!(segments.data_modes.len(), data_segments.len(), "native data modes must align with payload indices");
            let mut out = Vec::new();
            write_leb128_u32(&mut out, data_segments.len() as u32);
            for (mode, payload) in segments.data_modes.iter().zip(data_segments) {
                match mode {
                    NativeSegmentMode::Passive => out.push(1),
                    NativeSegmentMode::Active { index, offset } => {
                        if *index == 0 { out.push(0); }
                        else { out.push(2); write_leb128_u32(&mut out, *index); }
                        let wide = module.memory_is_64.get(*index as usize).copied().unwrap_or(false);
                        encode_segment_integer_expr(&mut out, offset, wide, module, ctx);
                        out.push(0x0b);
                    }
                    NativeSegmentMode::Declarative => panic!("data segments cannot be declarative"),
                }
                write_leb128_u32(&mut out, payload.len() as u32);
                out.extend_from_slice(payload);
            }
            return out;
        }
    }
    let Some(module) = chunks.first().filter(|chunk| !chunk.active_data_segments.is_empty()) else {
        return encode_data_section(data_segments);
    };
    let mut active = vec![None; data_segments.len()];
    for initializer in &module.active_data_segments {
        let entry = active.get_mut(initializer.data_index as usize)
            .expect("active data segment index must name a retained payload");
        assert!(entry.is_none(), "data segment has multiple initialization modes");
        *entry = Some(initializer);
    }
    let mut out = Vec::new();
    write_leb128_u32(&mut out, data_segments.len() as u32);
    for (index, payload) in data_segments.iter().enumerate() {
        if let Some(initializer) = active[index] {
            if initializer.memory_index == 0 {
                out.push(0);
            } else {
                out.push(2);
                write_leb128_u32(&mut out, initializer.memory_index);
            }
            if module.memory_is_64.get(initializer.memory_index as usize).copied().unwrap_or(false) {
                out.push(0x42);
                write_leb128_i64(&mut out, initializer.offset as i64);
            } else {
                assert!(u32::try_from(initializer.offset).is_ok(), "memory32 offset exceeds its address width");
                out.push(0x41);
                write_leb128_i32(&mut out, initializer.offset as i32);
            }
            out.push(0x0b);
        } else {
            out.push(1);
        }
        write_leb128_u32(&mut out, payload.len() as u32);
        out.extend_from_slice(payload);
    }
    out
}

// Offsets have a declared integer type, determined by their memory/table.
// Preserve integer bits and wrapping arithmetic; never evaluate an imported
// global during emission or infer the width from its current runtime value.
fn encode_segment_integer_expr(
    out: &mut Vec<u8>,
    expression: &vybe_runtime::chunk::ConstExpr,
    wide: bool,
    module: &Chunk,
    ctx: &super::types::WasmTypeContext,
) {
    use vybe_runtime::{Value, chunk::ConstExpr};
    match expression {
        ConstExpr::Value(Value::I32(value)) if !wide => {
            out.push(0x41); write_leb128_i32(out, *value);
        }
        ConstExpr::Value(Value::I64(value)) if wide => {
            out.push(0x42); write_leb128_i64(out, *value);
        }
        ConstExpr::GlobalGet(name) => {
            let old = module.globals.iter().position(|global| global == name)
                .expect("segment global must name a retained declaration");
            let index = *ctx.callables.global_map.get(old)
                .expect("segment global must have a mapped binary index");
            out.push(0x23); write_leb128_u32(out, index);
        }
        ConstExpr::Add(left, right) | ConstExpr::Mul(left, right) => {
            encode_segment_integer_expr(out, left, wide, module, ctx);
            encode_segment_integer_expr(out, right, wide, module, ctx);
            out.push(match (wide, expression) {
                (false, ConstExpr::Add(..)) => 0x6a,
                (false, _) => 0x6c,
                (true, ConstExpr::Add(..)) => 0x7c,
                (true, _) => 0x7e,
            });
        }
        _ => panic!("segment offset must be a constant expression of its declared integer type"),
    }
}

/// Emit a call to an imported function by (module, name) key.
pub fn emit_import_call(
    body: &mut Vec<u8>,
    rt_idx: &std::collections::HashMap<(&str, &str), usize>,
    module: &str,
    name: &str,
) {
    if let Some(&idx) = rt_idx.get(&(module, name)) {
        body.push(0x10); // call
        write_leb128_u32(body, idx as u32);
    } else {
        body.push(0x01); // nop
    }
}

/// Emit a call to vybe:rt runtime function (convenience).
pub fn emit_rt_call(
    body: &mut Vec<u8>,
    rt_idx: &std::collections::HashMap<(&str, &str), usize>,
    name: &str,
) {
    emit_import_call(body, rt_idx, "vybe:rt", name);
}

/// Box an i32 on the stack into externref via wasm:js-number fromI32.
pub fn emit_box_i32(body: &mut Vec<u8>, rt_idx: &std::collections::HashMap<(&str, &str), usize>) {
    emit_import_call(body, rt_idx, "wasm:js-number", "fromI32");
}

/// Box an f64 on the stack into externref via wasm:js-number fromF64.
pub fn emit_box_f64(body: &mut Vec<u8>, rt_idx: &std::collections::HashMap<(&str, &str), usize>) {
    emit_import_call(body, rt_idx, "wasm:js-number", "fromF64");
}

/// Unbox externref to f64 via wasm:js-number toF64.
pub fn emit_unbox_f64(body: &mut Vec<u8>, rt_idx: &std::collections::HashMap<(&str, &str), usize>) {
    emit_import_call(body, rt_idx, "wasm:js-number", "toF64");
}

/// Unbox externref to i32 via wasm:js-number toI32.
pub fn emit_unbox_i32(body: &mut Vec<u8>, rt_idx: &std::collections::HashMap<(&str, &str), usize>) {
    emit_import_call(body, rt_idx, "wasm:js-number", "toI32");
}

/// Box an i32 interpreted as unsigned into externref via
/// `wasm:js-number.fromU32`. JS sees a non-negative Number.
pub fn emit_box_u32(body: &mut Vec<u8>, rt_idx: &std::collections::HashMap<(&str, &str), usize>) {
    emit_import_call(body, rt_idx, "wasm:js-number", "fromU32");
}

/// Unbox externref to i32 **interpreted as unsigned** via
/// `wasm:js-number.toU32`. Used for compiler-emitted unsigned math.
pub fn emit_unbox_u32(body: &mut Vec<u8>, rt_idx: &std::collections::HashMap<(&str, &str), usize>) {
    emit_import_call(body, rt_idx, "wasm:js-number", "toU32");
}

/// Narrow numeric type test via `wasm:js-number.testI32`. Leaves i32 on
/// the stack (1 if the value is a Number representable as i32, else 0).
pub fn emit_test_i32(body: &mut Vec<u8>, rt_idx: &std::collections::HashMap<(&str, &str), usize>) {
    emit_import_call(body, rt_idx, "wasm:js-number", "testI32");
}

/// Same as `testI32` but checks non-negative u32 range.
pub fn emit_test_u32(body: &mut Vec<u8>, rt_idx: &std::collections::HashMap<(&str, &str), usize>) {
    emit_import_call(body, rt_idx, "wasm:js-number", "testU32");
}

/// Unbox externref to i32 via `wasm:js-boolean.cast`. Valid only when
/// the value has already tested as a JS boolean — host traps otherwise.
pub fn emit_unbox_bool(
    body: &mut Vec<u8>,
    rt_idx: &std::collections::HashMap<(&str, &str), usize>,
) {
    emit_import_call(body, rt_idx, "wasm:js-boolean", "cast");
}

/// Symbol identity check — `wasm:js-symbol.equals`. Consumes two
/// externrefs, produces i32 (1 if same symbol, 0 otherwise). Traps if
/// either operand isn't a symbol.
pub fn emit_symbol_equals(
    body: &mut Vec<u8>,
    rt_idx: &std::collections::HashMap<(&str, &str), usize>,
) {
    emit_import_call(body, rt_idx, "wasm:js-symbol", "equals");
}

pub fn encode_global_section_with_callables(
    globals: &[String],
    ctx: &super::types::WasmTypeContext,
    chunks: &[Chunk],
) -> Vec<u8> {
    let original = encode_global_section_with_descriptors(globals, Some(ctx));
    let mut at = 0;
    let count = vybe_runtime::opcode::read_leb_u32(&original, &mut at);
    let mut out = Vec::new();
    write_leb128_u32(&mut out, count + chunks.len() as u32 + 1);
    out.extend_from_slice(&original[at..]);
    for (i, c) in chunks.iter().enumerate() {
        out.push(0x64);
        write_leb128_i64(&mut out, ctx.callables.carrier as i64);
        out.push(0);
        super::callable::carrier(
            &mut out,
            &ctx.callables,
            i as u32,
            |b| super::callable::null(b, ctx.callables.captures as i32),
            ctx.callables.chunks[i],
        );
        out.push(0x0b);
    }
    out.extend([TYPE_I32, 1, 0x41, 0, 0x0b]); // once-per-instance link initialization flag
    out
}

pub fn encode_func_section_with_generators(
    chunks: &[Chunk],
    imports: usize,
    base: u32,
    resume: Option<u32>,
    ctx: &super::types::WasmTypeContext,
) -> Vec<u8> {
    let original = encode_func_section_with_native_resume(chunks, imports, base, resume);
    let mut ip = 0;
    let count = vybe_runtime::opcode::read_leb_u32(&original, &mut ip);
    let mut out = Vec::new();
    write_leb128_u32(&mut out, count + ctx.generators.len() as u32
        + u32::from(ctx.native_carrier_resume_function_idx.is_some())
        + u32::from(ctx.source_entry.is_some()));
    out.extend(&original[ip..]);
    if ctx.native_carrier_resume_function_idx.is_some(){write_leb128_u32(&mut out,ctx.native_carrier_resume_type_idx);}
    for e in &ctx.generators {
        write_leb128_u32(&mut out, ctx.callables.calls[&(e.user_arity, 1)]);
    }
    if let Some((_, _, type_index, _)) = ctx.source_entry {
        write_leb128_u32(&mut out, type_index);
    }
    out
}
pub fn encode_element_section_with_generators(
    chunks: &[Chunk],
    _imports: usize,
    ctx: &super::types::WasmTypeContext,
) -> Vec<u8> {
    // Preserve the original element form/table chunk ordinals; map only real function references.
    let mut out = Vec::new();
    let native = chunks.first().and_then(|module| module.native_segments.as_ref());
    let original_count = native.map_or(0, |segments| segments.elements.len());
    write_leb128_u32(&mut out, original_count as u32 + 1 + u32::from(!ctx.generators.is_empty() || ctx.native_resume_function_idx.is_some()));
    // Keep original indices: table.init/elem.drop address these segments,
    // whereas the callable and generator segments are writer-owned suffixes.
    if let Some(native) = native {
        let module = &chunks[0];
        for segment in &native.elements {
            use vybe_runtime::chunk::NativeSegmentMode;
            match &segment.mode {
                NativeSegmentMode::Passive => out.push(5),
                NativeSegmentMode::Declarative => out.push(7),
                NativeSegmentMode::Active { index, offset } => {
                    out.push(6);
                    write_leb128_u32(&mut out, *index);
                    let wide = module.table_is_64.get(*index as usize).copied().unwrap_or(false);
                    encode_segment_integer_expr(&mut out, offset, wide, module, ctx);
                    out.push(0x0b);
                }
            }
            encode_segment_reference_type(&mut out, segment.reference_type, ctx);
            write_leb128_u32(&mut out, segment.items.len() as u32);
            for item in segment.items.iter() {
                encode_segment_reference_expr(&mut out, item, segment.reference_type, module, ctx);
                out.push(0x0b);
            }
        }
    }
    if ctx.internal_callable_table == 0 {
        write_leb128_u32(&mut out, 0);
    } else {
        write_leb128_u32(&mut out, 2);
        write_leb128_u32(&mut out, ctx.internal_callable_table);
    }
    out.extend([0x41, 0, 0x0b]);
    if ctx.internal_callable_table != 0 { out.push(0); }
    write_leb128_u32(&mut out, chunks.len() as u32);
    for target in &ctx.callables.targets {
        write_leb128_u32(&mut out, *target);
    }
    if !ctx.generators.is_empty() || ctx.native_resume_function_idx.is_some() {
        write_leb128_u32(&mut out, 3);
        out.push(0);
        write_leb128_u32(&mut out, ctx.generators.len() as u32 + u32::from(ctx.native_resume_function_idx.is_some()));
        if let Some(driver)=ctx.native_resume_function_idx{write_leb128_u32(&mut out,driver);}
        for e in &ctx.generators {
            write_leb128_u32(&mut out, e.body_func);
        }
    }
    out
}

fn encode_segment_reference_type(
    out: &mut Vec<u8>,
    reference: (vybe_runtime::opcode::heaptype::HeapType, bool, bool, bool),
    ctx: &super::types::WasmTypeContext,
) {
    use vybe_runtime::opcode::heaptype::HeapType;
    let (heap, nullable, exact, shared) = reference;
    match heap {
        HeapType::Abstract(byte) if nullable && !shared => out.push(byte),
        HeapType::Abstract(byte) => {
            out.push(if nullable { 0x63 } else { 0x64 });
            if shared { out.push(0x65); }
            out.push(byte);
        }
        HeapType::Concrete(module_index) => {
            out.push(if nullable { 0x63 } else { 0x64 });
            let index = ctx.struct_type_by_index(module_index)
                .expect("segment reference type must resolve in the module type space");
            if exact { write_exact_heaptype(out, index); }
            else { write_leb128_i64(out, index as i64); }
        }
    }
}

fn encode_segment_reference_expr(
    out: &mut Vec<u8>,
    expression: &vybe_runtime::chunk::ConstExpr,
    reference: (vybe_runtime::opcode::heaptype::HeapType, bool, bool, bool),
    module: &Chunk,
    ctx: &super::types::WasmTypeContext,
) {
    use vybe_runtime::{Value, chunk::ConstExpr, opcode::heaptype::HeapType};
    match expression {
        ConstExpr::RefFunc(chunk) => {
            out.push(0xd2);
            write_leb128_u32(out, *ctx.callables.targets.get(*chunk)
                .expect("element function must resolve to a retained chunk"));
        }
        ConstExpr::Value(Value::Null) => {
            out.push(0xd0);
            match reference.0 {
                HeapType::Abstract(byte) => {
                    if reference.3 { out.push(0x65); }
                    out.push(byte);
                }
                HeapType::Concrete(module_index) => {
                    let index = ctx.struct_type_by_index(module_index)
                        .expect("element null heap type must resolve");
                    if reference.2 { write_exact_heaptype(out, index); }
                    else { write_leb128_i64(out, index as i64); }
                }
            }
        }
        ConstExpr::GlobalGet(name) => {
            let old = module.globals.iter().position(|global| global == name)
                .expect("element global must name a retained declaration");
            out.push(0x23);
            write_leb128_u32(out, ctx.callables.global_map[old]);
        }
        ConstExpr::RefI31(value) => {
            encode_segment_integer_expr(out, value, false, module, ctx);
            out.extend_from_slice(&[0xfb, 0x1c]);
        }
        // Imported function identities and GC constructors require their
        // retained native declaration graph, not Source arity/name inference.
        // Refuse unsupported lowering rather than silently losing payloads.
        _ => panic!("native element initializer requires retained native declaration lowering"),
    }
}
