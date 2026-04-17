//! WASM module section encoding (import, function, memory, export).

use super::encoding::*;
use crate::Chunk;
use crate::opcode::Op;

/// Collect all runtime imports needed by the chunks.
/// Returns (module, name) pairs for both vybe:rt and wasm:js-* builtins.
pub fn collect_rt_imports(chunks: &[Chunk]) -> Vec<(&'static str, &'static str)> {
    let mut needed: Vec<(&str, &str)> = Vec::new();
    let mut seen = std::collections::HashSet::new();

    // Always include js-number builtins for boxing/unboxing in .wasm output
    let js_builtins: &[(&str, &str)] = &[
        ("wasm:js-number", "fromF64"),
        ("wasm:js-number", "fromI32"),
        ("wasm:js-number", "toF64"),
        ("wasm:js-number", "toI32"),
        ("wasm:js-string", "test"),
        ("wasm:js-string", "concat"),
        ("wasm:js-boolean", "test"),
        ("wasm:js-undefined", "test"),
    ];
    for &(module, name) in js_builtins {
        let key = (module, name);
        if seen.insert(key) { needed.push(key); }
    }

    // Scan chunks for dynamic ops that need vybe:rt runtime calls
    for chunk in chunks {
        let mut ip = 0;
        while ip < chunk.code.len() {
            if ip + 1 >= chunk.code.len() { break; }
            if let Some(op) = Op::decode(chunk.code[ip], chunk.code[ip + 1]) {
                let rt_name: Option<&str> = if op == Op::DYN_ADD { Some("dyn_add") }
                    else if op == Op::DYN_EQ { Some("dyn_eq") }
                    else if op == Op::DYN_NE { Some("dyn_ne") }
                    else if op == Op::DYN_LT { Some("dyn_lt") }
                    else if op == Op::DYN_GT { Some("dyn_gt") }
                    else if op == Op::DYN_LE { Some("dyn_le") }
                    else if op == Op::DYN_GE { Some("dyn_ge") }
                    else if op == Op::DYN_NEG { Some("dyn_neg") }
                    else if op == Op::DYN_NOT { Some("dyn_not") }
                    else if op == Op::DYN_TO_BOOL { Some("dyn_to_bool") }
                    else if op == Op::STR_CONCAT { Some("str_concat") }
                    else { None };
                if let Some(name) = rt_name {
                    let key = ("vybe:rt", name);
                    if seen.insert(key) { needed.push(key); }
                }
                ip += super::code::opcode_size(op, &chunk.code, ip);
            } else {
                ip += 2;
            }
        }
    }
    needed
}

pub fn encode_import_section(chunks: &[Chunk], rt_imports: &[(&str, &str)], func_type_base: u32) -> Vec<u8> {
    let mut out = Vec::new();
    let host_imports = chunks.first().map(|c| c.imports.len()).unwrap_or(0);
    let total = host_imports + rt_imports.len();
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

    // Runtime + builtin imports (mixed modules: vybe:rt, wasm:js-number, etc.)
    for (i, (module, name)) in rt_imports.iter().enumerate() {
        write_name(&mut out, module);
        write_name(&mut out, name);
        out.push(0x00); // func import
        write_leb128_u32(&mut out, func_type_base + (host_imports + i) as u32);
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
    let mut out = Vec::new();
    write_leb128_u32(&mut out, 1);
    out.push(0x00); // no max
    write_leb128_u32(&mut out, 1); // 1 page
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

/// Emit a call to an imported function by (module, name) key.
pub fn emit_import_call(body: &mut Vec<u8>, rt_idx: &std::collections::HashMap<(&str, &str), usize>, module: &str, name: &str) {
    if let Some(&idx) = rt_idx.get(&(module, name)) {
        body.push(0x10); // call
        write_leb128_u32(body, idx as u32);
    } else {
        body.push(0x01); // nop
    }
}

/// Emit a call to vybe:rt runtime function (convenience).
pub fn emit_rt_call(body: &mut Vec<u8>, rt_idx: &std::collections::HashMap<(&str, &str), usize>, name: &str) {
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
