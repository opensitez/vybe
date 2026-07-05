//! WASM module section encoding (import, function, memory, export).

use crate::encoding::*;
use vybe_bytecode::Chunk;
use vybe_bytecode::opcode::Op;

/// Collect all runtime imports needed by the chunks.
/// Returns (module, name) pairs for both vybe:rt and wasm:js-* builtins.
pub fn collect_rt_imports(_chunks: &[Chunk]) -> Vec<(&'static str, &'static str)> {
    let mut needed: Vec<(&str, &str)> = Vec::new();
    let mut seen = std::collections::HashSet::new();

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

    // No vybe:rt imports — all ops lower to inline WASM + standard wasm:js-* builtins
    needed
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
    rt_imports: &[(&str, &str)],
    func_type_base: u32,
) -> Vec<u8> {
    let mut out = Vec::new();
    let host_imports = chunks.first().map(|c| c.imports.len()).unwrap_or(0);
    let globals = rt_globals();
    let total = host_imports + rt_imports.len() + globals.len();
    write_leb128_u32(&mut out, total as u32);

    // Host imports from chunk 0
    if let Some(chunk) = chunks.first() {
        for (i, import) in chunk.imports.iter().enumerate() {
            write_name(&mut out, &import.module);
            write_name(&mut out, &import.name);
            out.push(0x00); // func import
            write_leb128_u32(&mut out, func_type_base + i as u32);
        }
    }

    // Runtime + builtin function imports (mixed modules: vybe:rt, wasm:js-*, …)
    for (i, (module, name)) in rt_imports.iter().enumerate() {
        write_name(&mut out, module);
        write_name(&mut out, name);
        out.push(0x00); // func import
        write_leb128_u32(&mut out, func_type_base + (host_imports + i) as u32);
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

pub fn encode_memory_section() -> Vec<u8> {
    // Default: single non-shared memory, 1 page min, no max.
    encode_memory_section_with(1, None, false)
}

pub fn module_uses_memory64(chunks: &[Chunk]) -> bool {
    for chunk in chunks {
        let mut ip = 0;
        while ip + 3 < chunk.code.len() {
            let g = ((chunk.code[ip] as u16) << 8) | chunk.code[ip + 1] as u16;
            let s = ((chunk.code[ip + 2] as u16) << 8) | chunk.code[ip + 3] as u16;
            let Some(op) = Op::decode(g, s) else {
                ip += 4;
                continue;
            };
            if op == Op::I64_MEMORY_SIZE
                || op == Op::I64_MEMORY_GROW
                || op == Op::I32_LOAD_64
                || op == Op::I64_LOAD_64
                || op == Op::F64_LOAD_64
                || op == Op::I32_STORE_64
                || op == Op::I64_STORE_64
                || op == Op::F64_STORE_64
            {
                return true;
            }
            ip += crate::writer::code::opcode_size(op, &chunk.code, ip);
        }
    }
    false
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

pub fn encode_export_section(_chunks: &[Chunk], import_count: usize) -> Vec<u8> {
    let mut out = Vec::new();
    write_leb128_u32(&mut out, 2); // export memory + main func
    // Memory
    write_name(&mut out, "memory");
    out.push(0x02); // memory export
    write_leb128_u32(&mut out, 0);
    // Main function (first chunk after imports)
    write_name(&mut out, "_start");
    out.push(0x00); // func export
    write_leb128_u32(&mut out, import_count as u32);
    out
}

/// Collect all unique global variable names from chunks and build name→index map.
pub fn collect_globals(chunks: &[Chunk]) -> (Vec<String>, std::collections::HashMap<String, u32>) {
    let mut globals = Vec::new();
    let mut global_map = std::collections::HashMap::new();

    for chunk in chunks {
        let mut ip = 0;
        while ip < chunk.code.len() {
            if ip + 3 >= chunk.code.len() {
                break;
            }
            let group = ((chunk.code[ip] as u16) << 8) | chunk.code[ip + 1] as u16;
            let sub = ((chunk.code[ip + 2] as u16) << 8) | chunk.code[ip + 3] as u16;
            if let Some(op) = Op::decode(group, sub) {
                if op == Op::GLOBAL_GET || op == Op::GLOBAL_SET {
                    let name_idx = ((chunk.code[ip + 4] as u16) << 8) | chunk.code[ip + 5] as u16;
                    if let Some(vybe_bytecode::value::Value::String(name)) =
                        chunk.constants.get(name_idx as usize)
                    {
                        let name_str = name.to_string();
                        if !global_map.contains_key(&name_str) {
                            let idx = globals.len() as u32;
                            global_map.insert(name_str.clone(), idx);
                            globals.push(name_str);
                        }
                    }
                }
                ip += crate::writer::code::opcode_size(op, &chunk.code, ip);
            } else {
                ip += 4;
            }
        }
    }
    (globals, global_map)
}

/// Encode the global section — all globals are (mut externref), initialized to null.
pub fn encode_global_section(globals: &[String]) -> Vec<u8> {
    let mut out = Vec::new();
    write_leb128_u32(&mut out, globals.len() as u32);
    for _ in globals {
        out.push(0x6F); // externref
        out.push(0x01); // mutable
        // Init expr: ref.null extern
        out.push(0xD0);
        out.push(0x6F); // ref.null extern
        out.push(0x0B); // end
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

/// String format of an i32 via `wasm:js-string.fromI32`. Consumes i32,
/// produces externref (a JS string). Mirrors `String(n)` in JS.
pub fn emit_str_from_i32(
    body: &mut Vec<u8>,
    rt_idx: &std::collections::HashMap<(&str, &str), usize>,
) {
    emit_import_call(body, rt_idx, "wasm:js-string", "fromI32");
}

/// Unsigned-i32 → string via `wasm:js-string.fromU32`.
pub fn emit_str_from_u32(
    body: &mut Vec<u8>,
    rt_idx: &std::collections::HashMap<(&str, &str), usize>,
) {
    emit_import_call(body, rt_idx, "wasm:js-string", "fromU32");
}

/// i64 → string via `wasm:js-string.fromI64`.
pub fn emit_str_from_i64(
    body: &mut Vec<u8>,
    rt_idx: &std::collections::HashMap<(&str, &str), usize>,
) {
    emit_import_call(body, rt_idx, "wasm:js-string", "fromI64");
}

/// Unsigned-i64 → string via `wasm:js-string.fromU64`.
pub fn emit_str_from_u64(
    body: &mut Vec<u8>,
    rt_idx: &std::collections::HashMap<(&str, &str), usize>,
) {
    emit_import_call(body, rt_idx, "wasm:js-string", "fromU64");
}

/// f64 → string via `wasm:js-string.fromF64`. Matches `String(n)` when
/// `n` is a finite JS Number.
pub fn emit_str_from_f64(
    body: &mut Vec<u8>,
    rt_idx: &std::collections::HashMap<(&str, &str), usize>,
) {
    emit_import_call(body, rt_idx, "wasm:js-string", "fromF64");
}

/// Validating string cast — `wasm:js-string.cast`. Traps if the value
/// isn't a string. Equivalent to `(stringref) <: anyref` in spec terms.
pub fn emit_str_cast(body: &mut Vec<u8>, rt_idx: &std::collections::HashMap<(&str, &str), usize>) {
    emit_import_call(body, rt_idx, "wasm:js-string", "cast");
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
