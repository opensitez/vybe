//! WASM binary format reader/writer.
//!
//! Produces valid WASM modules that tools like wasm-dis can parse.
//! Dynamic operations (dyn_add, struct_get, globals) go through
//! imported vybe:rt functions. Standard WASM ops use real opcodes.

use crate::{Chunk, Op};
use crate::value::Value;
use std::sync::Arc;

const WASM_MAGIC: [u8; 4] = [0x00, 0x61, 0x73, 0x6D];
const WASM_VERSION: [u8; 4] = [0x01, 0x00, 0x00, 0x00];

const SECTION_CUSTOM: u8 = 0;
const SECTION_TYPE: u8 = 1;
const SECTION_IMPORT: u8 = 2;
const SECTION_FUNCTION: u8 = 3;
const SECTION_MEMORY: u8 = 5;
const SECTION_EXPORT: u8 = 7;
const SECTION_CODE: u8 = 10;

const TYPE_FUNC: u8 = 0x60;
const _TYPE_I32: u8 = 0x7F;
const TYPE_I64: u8 = 0x7E;
const _TYPE_F64: u8 = 0x7C;
const TYPE_VOID: u8 = 0x40;

// ============================================================
// Writer
// ============================================================

pub fn write_wasm(chunks: &[Chunk]) -> Vec<u8> {
    let mut out = Vec::new();
    out.extend_from_slice(&WASM_MAGIC);
    out.extend_from_slice(&WASM_VERSION);

    // Custom section: Vybe metadata for round-trip
    write_section(&mut out, SECTION_CUSTOM, &encode_custom_section(chunks));

    // Collect all imports (host + runtime)
    let rt_imports = collect_rt_imports(chunks);
    let total_imports = rt_imports.len();

    // Type section: one type for each unique signature
    // Simplified: type 0 = () -> (), type 1 = (i32) -> (), etc.
    write_section(&mut out, SECTION_TYPE, &encode_type_section(chunks, total_imports));

    // Import section: host imports from chunks + vybe:rt runtime
    write_section(&mut out, SECTION_IMPORT, &encode_import_section(chunks, &rt_imports));

    // Function section
    write_section(&mut out, SECTION_FUNCTION, &encode_func_section(chunks, total_imports));

    // Memory section
    write_section(&mut out, SECTION_MEMORY, &encode_memory_section());

    // Export section
    write_section(&mut out, SECTION_EXPORT, &encode_export_section(chunks, total_imports));

    // Code section with real WASM opcodes
    write_section(&mut out, SECTION_CODE, &encode_code_section(chunks, &rt_imports));

    out
}

/// Collect all vybe:rt function names needed by the chunks
fn collect_rt_imports(chunks: &[Chunk]) -> Vec<(&'static str, &'static str)> {
    let mut needed = std::collections::HashSet::new();
    for chunk in chunks {
        let mut ip = 0;
        while ip < chunk.code.len() {
            if let Some(op) = Op::decode(chunk.code[ip], chunk.code[ip + 1]) {
                match op {
                    _ if op == Op::DYN_ADD => { needed.insert("dyn_add"); }
                    _ if op == Op::DYN_EQ => { needed.insert("dyn_eq"); }
                    _ if op == Op::DYN_NE => { needed.insert("dyn_ne"); }
                    _ if op == Op::DYN_LT => { needed.insert("dyn_lt"); }
                    _ if op == Op::DYN_GT => { needed.insert("dyn_gt"); }
                    _ if op == Op::DYN_LE => { needed.insert("dyn_le"); }
                    _ if op == Op::DYN_GE => { needed.insert("dyn_ge"); }
                    _ if op == Op::DYN_NEG => { needed.insert("dyn_neg"); }
                    _ if op == Op::DYN_NOT => { needed.insert("dyn_not"); }
                    _ if op == Op::DYN_TO_BOOL => { needed.insert("dyn_to_bool"); }
                    _ if op == Op::STR_CONCAT => { needed.insert("str_concat"); }
                    _ if op == Op::STRUCT_GET => { needed.insert("get_prop"); }
                    _ if op == Op::STRUCT_SET => { needed.insert("set_prop"); }
                    _ if op == Op::STRUCT_NEW => { needed.insert("new_object"); }
                    _ if op == Op::ARRAY_GET => { needed.insert("array_get"); }
                    _ if op == Op::ARRAY_SET => { needed.insert("array_set"); }
                    _ if op == Op::ARRAY_NEW => { needed.insert("new_array"); }
                    _ if op == Op::GLOBAL_GET => { needed.insert("global_get"); }
                    _ if op == Op::GLOBAL_SET => { needed.insert("global_set"); }
                    _ => {}
                }
                ip += opcode_size(op, &chunk.code, ip);
            } else {
                ip += 2; // skip unknown 2-byte opcode
            }
        }
    }
    // Return in deterministic order
    let mut result: Vec<(&str, &str)> = needed.into_iter()
        .map(|name| ("vybe:rt", name))
        .collect();
    result.sort_by_key(|&(_, n)| n);
    result
}

fn encode_type_section(chunks: &[Chunk], import_count: usize) -> Vec<u8> {
    let mut out = Vec::new();
    let count = import_count + chunks.len();
    write_leb128_u32(&mut out, count as u32);
    // All import functions: variadic → simplified as () -> i64
    for _ in 0..import_count {
        out.push(TYPE_FUNC);
        write_leb128_u32(&mut out, 0); // 0 params
        write_leb128_u32(&mut out, 1); // 1 result
        out.push(TYPE_I64);
    }
    // Chunk functions
    for chunk in chunks {
        out.push(TYPE_FUNC);
        write_leb128_u32(&mut out, chunk.arity as u32);
        for _ in 0..chunk.arity { out.push(TYPE_I64); }
        write_leb128_u32(&mut out, 1);
        out.push(TYPE_I64);
    }
    out
}

fn encode_import_section(chunks: &[Chunk], rt_imports: &[(&str, &str)]) -> Vec<u8> {
    let mut out = Vec::new();
    let host_imports = chunks.first().map(|c| c.imports.len()).unwrap_or(0);
    write_leb128_u32(&mut out, (host_imports + rt_imports.len()) as u32);

    // Host imports from chunk 0
    if let Some(chunk) = chunks.first() {
        for (i, import) in chunk.imports.iter().enumerate() {
            write_name(&mut out, &import.module);
            write_name(&mut out, &import.name);
            out.push(0x00); // func import
            write_leb128_u32(&mut out, i as u32); // type index
        }
    }

    // Runtime imports
    for (i, (module, name)) in rt_imports.iter().enumerate() {
        write_name(&mut out, module);
        write_name(&mut out, name);
        out.push(0x00);
        write_leb128_u32(&mut out, (host_imports + i) as u32);
    }
    out
}

fn encode_func_section(chunks: &[Chunk], import_count: usize) -> Vec<u8> {
    let mut out = Vec::new();
    write_leb128_u32(&mut out, chunks.len() as u32);
    for (i, _) in chunks.iter().enumerate() {
        write_leb128_u32(&mut out, (import_count + i) as u32);
    }
    out
}

fn encode_memory_section() -> Vec<u8> {
    let mut out = Vec::new();
    write_leb128_u32(&mut out, 1);
    out.push(0x00); // no max
    write_leb128_u32(&mut out, 1); // 1 page
    out
}

fn encode_export_section(_chunks: &[Chunk], import_count: usize) -> Vec<u8> {
    let mut out = Vec::new();
    write_leb128_u32(&mut out, 1);
    write_name(&mut out, "_start");
    out.push(0x00); // func export
    write_leb128_u32(&mut out, import_count as u32); // first non-import func
    out
}

fn encode_code_section(chunks: &[Chunk], rt_imports: &[(&str, &str)]) -> Vec<u8> {
    let host_import_count = chunks.first().map(|c| c.imports.len()).unwrap_or(0);

    // Build rt import name → index map
    let mut rt_idx: std::collections::HashMap<&str, usize> = std::collections::HashMap::new();
    for (i, (_, name)) in rt_imports.iter().enumerate() {
        rt_idx.insert(name, host_import_count + i);
    }

    let mut out = Vec::new();
    write_leb128_u32(&mut out, chunks.len() as u32);

    for chunk in chunks {
        let mut body = Vec::new();

        // Locals
        if chunk.local_count > 0 {
            write_leb128_u32(&mut body, 1);
            write_leb128_u32(&mut body, chunk.local_count as u32);
            body.push(TYPE_I64);
        } else {
            write_leb128_u32(&mut body, 0);
        }

        // Translate opcodes
        let mut ip = 0;
        while ip < chunk.code.len() {
            if ip + 1 >= chunk.code.len() { break; }
            let op = match Op::decode(chunk.code[ip], chunk.code[ip + 1]) {
                Some(op) => op,
                None => { ip += 2; continue; }
            };
            ip += 2; // skip past 2-byte opcode

            match op {
                // --- Standard WASM ops ---
                _ if op == Op::LOCAL_GET => { body.push(0x20); write_leb128_u32(&mut body, read_u16(&chunk.code, &mut ip) as u32); }
                _ if op == Op::LOCAL_SET => { body.push(0x22); write_leb128_u32(&mut body, read_u16(&chunk.code, &mut ip) as u32); } // tee
                _ if op == Op::DROP => body.push(0x1A),
                _ if op == Op::RETURN => body.push(0x0F),
                _ if op == Op::HALT => body.push(0x00), // unreachable

                _ if op == Op::F64_ADD => body.push(0xA0),
                _ if op == Op::F64_SUB => body.push(0xA1),
                _ if op == Op::F64_MUL => body.push(0xA2),
                _ if op == Op::F64_DIV => body.push(0xA3),
                _ if op == Op::I32_ADD => body.push(0x6A),
                _ if op == Op::I32_SUB => body.push(0x6B),
                _ if op == Op::I32_MUL => body.push(0x6C),
                _ if op == Op::I32_DIV_S => body.push(0x6D),
                _ if op == Op::I32_DIV_U => body.push(0x6E),
                _ if op == Op::I32_REM_S => body.push(0x6F),
                _ if op == Op::I32_REM_U => body.push(0x70),
                _ if op == Op::I32_AND => body.push(0x71),
                _ if op == Op::I32_OR => body.push(0x72),
                _ if op == Op::I32_XOR => body.push(0x73),
                _ if op == Op::I32_SHL => body.push(0x74),
                _ if op == Op::I32_SHR_S => body.push(0x75),
                _ if op == Op::I32_SHR_U => body.push(0x76),
                _ if op == Op::I32_ROTL => body.push(0x77),
                _ if op == Op::I32_ROTR => body.push(0x78),
                _ if op == Op::I32_CLZ => body.push(0x67),
                _ if op == Op::I32_CTZ => body.push(0x68),
                _ if op == Op::I32_POPCNT => body.push(0x69),
                _ if op == Op::I32_EQZ => body.push(0x45),

                // i64
                _ if op == Op::I64_ADD => body.push(0x7C),
                _ if op == Op::I64_SUB => body.push(0x7D),
                _ if op == Op::I64_MUL => body.push(0x7E),
                _ if op == Op::I64_DIV_S => body.push(0x7F),
                _ if op == Op::I64_DIV_U => body.push(0x80),
                _ if op == Op::I64_REM_S => body.push(0x81),
                _ if op == Op::I64_REM_U => body.push(0x82),
                _ if op == Op::I64_AND => body.push(0x83),
                _ if op == Op::I64_OR => body.push(0x84),
                _ if op == Op::I64_XOR => body.push(0x85),
                _ if op == Op::I64_SHL => body.push(0x86),
                _ if op == Op::I64_SHR_S => body.push(0x87),
                _ if op == Op::I64_SHR_U => body.push(0x88),
                _ if op == Op::I64_ROTL => body.push(0x89),
                _ if op == Op::I64_ROTR => body.push(0x8A),
                _ if op == Op::I64_CLZ => body.push(0x79),
                _ if op == Op::I64_CTZ => body.push(0x7A),
                _ if op == Op::I64_POPCNT => body.push(0x7B),
                _ if op == Op::I64_EQZ => body.push(0x50),

                // f64 math
                _ if op == Op::F64_ABS => body.push(0x99),
                _ if op == Op::F64_NEG => body.push(0x9A),
                _ if op == Op::F64_CEIL => body.push(0x9B),
                _ if op == Op::F64_FLOOR => body.push(0x9C),
                _ if op == Op::F64_TRUNC => body.push(0x9D),
                _ if op == Op::F64_NEAREST => body.push(0x9E),
                _ if op == Op::F64_SQRT => body.push(0x9F),
                _ if op == Op::F64_MIN => body.push(0xA4),
                _ if op == Op::F64_MAX => body.push(0xA5),
                _ if op == Op::F64_COPYSIGN => body.push(0xA6),

                // f32 math
                _ if op == Op::F32_ABS => body.push(0x8B),
                _ if op == Op::F32_NEG => body.push(0x8C),
                _ if op == Op::F32_CEIL => body.push(0x8D),
                _ if op == Op::F32_FLOOR => body.push(0x8E),
                _ if op == Op::F32_TRUNC => body.push(0x8F),
                _ if op == Op::F32_NEAREST => body.push(0x90),
                _ if op == Op::F32_SQRT => body.push(0x91),
                _ if op == Op::F32_MIN => body.push(0x96),
                _ if op == Op::F32_MAX => body.push(0x97),
                _ if op == Op::F32_COPYSIGN => body.push(0x98),

                // select
                _ if op == Op::SELECT => body.push(0x1B),

                // conversions
                _ if op == Op::I32_FROM_F64 => body.push(0xAA),
                _ if op == Op::F64_FROM_I32 => body.push(0xB7),
                _ if op == Op::I32_WRAP_I64 => body.push(0xA7),
                _ if op == Op::I64_EXTEND_I32_S => body.push(0xAC),
                _ if op == Op::I64_EXTEND_I32_U => body.push(0xAD),
                _ if op == Op::I64_TRUNC_F64_S => body.push(0xB0),
                _ if op == Op::I64_TRUNC_F64_U => body.push(0xB1),
                _ if op == Op::F64_PROMOTE_F32 => body.push(0xB9),
                _ if op == Op::F32_DEMOTE_F64 => body.push(0xB6),
                _ if op == Op::I32_REINTERPRET_F32 => body.push(0xBC),
                _ if op == Op::I64_REINTERPRET_F64 => body.push(0xBD),
                _ if op == Op::F32_REINTERPRET_I32 => body.push(0xBA),
                _ if op == Op::F64_REINTERPRET_I64 => body.push(0xBB),
                _ if op == Op::I32_EXTEND8_S => body.push(0xC0),
                _ if op == Op::I32_EXTEND16_S => body.push(0xC1),
                _ if op == Op::I64_EXTEND8_S => body.push(0xC2),
                _ if op == Op::I64_EXTEND16_S => body.push(0xC3),
                _ if op == Op::I64_EXTEND32_S => body.push(0xC4),

                _ if op == Op::MEMORY_SIZE => { body.push(0x3F); body.push(0x00); }
                _ if op == Op::MEMORY_GROW => { body.push(0x40); body.push(0x00); }
                _ if op == Op::I32_LOAD => { body.push(0x28); body.push(0x02); body.push(0x00); }
                _ if op == Op::I64_LOAD => { body.push(0x29); body.push(0x03); body.push(0x00); }
                _ if op == Op::F32_LOAD => { body.push(0x2A); body.push(0x02); body.push(0x00); }
                _ if op == Op::F64_LOAD => { body.push(0x2B); body.push(0x03); body.push(0x00); }
                _ if op == Op::I32_LOAD8_S => { body.push(0x2C); body.push(0x00); body.push(0x00); }
                _ if op == Op::I32_LOAD8_U => { body.push(0x2D); body.push(0x00); body.push(0x00); }
                _ if op == Op::I32_LOAD16_S => { body.push(0x2E); body.push(0x01); body.push(0x00); }
                _ if op == Op::I32_LOAD16_U => { body.push(0x2F); body.push(0x01); body.push(0x00); }
                _ if op == Op::I64_LOAD8_S => { body.push(0x30); body.push(0x00); body.push(0x00); }
                _ if op == Op::I64_LOAD8_U => { body.push(0x31); body.push(0x00); body.push(0x00); }
                _ if op == Op::I64_LOAD16_S => { body.push(0x32); body.push(0x01); body.push(0x00); }
                _ if op == Op::I64_LOAD16_U => { body.push(0x33); body.push(0x01); body.push(0x00); }
                _ if op == Op::I64_LOAD32_S => { body.push(0x34); body.push(0x02); body.push(0x00); }
                _ if op == Op::I64_LOAD32_U => { body.push(0x35); body.push(0x02); body.push(0x00); }
                _ if op == Op::I32_STORE => { body.push(0x36); body.push(0x02); body.push(0x00); }
                _ if op == Op::I64_STORE => { body.push(0x37); body.push(0x03); body.push(0x00); }
                _ if op == Op::F32_STORE => { body.push(0x38); body.push(0x02); body.push(0x00); }
                _ if op == Op::F64_STORE => { body.push(0x39); body.push(0x03); body.push(0x00); }
                _ if op == Op::I32_STORE8 => { body.push(0x3A); body.push(0x00); body.push(0x00); }
                _ if op == Op::I32_STORE16 => { body.push(0x3B); body.push(0x01); body.push(0x00); }
                _ if op == Op::I64_STORE8 => { body.push(0x3C); body.push(0x00); body.push(0x00); }
                _ if op == Op::I64_STORE16 => { body.push(0x3D); body.push(0x01); body.push(0x00); }
                _ if op == Op::I64_STORE32 => { body.push(0x3E); body.push(0x02); body.push(0x00); }

                _ if op == Op::NULL => { body.push(0x42); write_leb128_i64(&mut body, 0); } // i64.const 0
                _ if op == Op::TRUE => { body.push(0x42); write_leb128_i64(&mut body, 1); }
                _ if op == Op::FALSE => { body.push(0x42); write_leb128_i64(&mut body, 0); }

                _ if op == Op::CONST => {
                    let idx = read_u16(&chunk.code, &mut ip);
                    if let Some(val) = chunk.constants.get(idx as usize) {
                        match val {
                            Value::F64(n) => { body.push(0x44); body.extend_from_slice(&n.to_le_bytes()); }
                            Value::I32(n) => { body.push(0x41); write_leb128_i32(&mut body, *n); }
                            Value::I64(n) => { body.push(0x42); write_leb128_i64(&mut body, *n); }
                            _ => { body.push(0x42); write_leb128_i64(&mut body, 0); } // placeholder
                        }
                    }
                }

                _ if op == Op::CALL => { body.push(0x10); let argc = chunk.code[ip]; ip += 1; write_leb128_u32(&mut body, argc as u32); }
                _ if op == Op::CALL_IMPORT => {
                    let import_idx = read_u16(&chunk.code, &mut ip);
                    let _argc = chunk.code[ip]; ip += 1;
                    body.push(0x10);
                    write_leb128_u32(&mut body, import_idx as u32);
                }

                _ if op == Op::BR => { let _ = read_i16(&chunk.code, &mut ip); body.push(0x0C); write_leb128_u32(&mut body, 0); }
                _ if op == Op::BR_IF_FALSE || op == Op::BR_IF_TRUE => { let _ = read_i16(&chunk.code, &mut ip); body.push(0x0D); write_leb128_u32(&mut body, 0); }
                _ if op == Op::BLOCK => { let _ = read_u16(&chunk.code, &mut ip); body.push(0x02); body.push(TYPE_VOID); }
                _ if op == Op::LOOP => { let _ = read_u16(&chunk.code, &mut ip); body.push(0x03); body.push(TYPE_VOID); }
                _ if op == Op::END => body.push(0x0B),

                // --- Dynamic ops → call vybe:rt import ---
                _ if op == Op::DYN_ADD => emit_rt_call(&mut body, &rt_idx, "dyn_add"),
                _ if op == Op::DYN_EQ => emit_rt_call(&mut body, &rt_idx, "dyn_eq"),
                _ if op == Op::DYN_NE => emit_rt_call(&mut body, &rt_idx, "dyn_ne"),
                _ if op == Op::DYN_LT => emit_rt_call(&mut body, &rt_idx, "dyn_lt"),
                _ if op == Op::DYN_GT => emit_rt_call(&mut body, &rt_idx, "dyn_gt"),
                _ if op == Op::DYN_LE => emit_rt_call(&mut body, &rt_idx, "dyn_le"),
                _ if op == Op::DYN_GE => emit_rt_call(&mut body, &rt_idx, "dyn_ge"),
                _ if op == Op::DYN_NEG => emit_rt_call(&mut body, &rt_idx, "dyn_neg"),
                _ if op == Op::DYN_NOT => emit_rt_call(&mut body, &rt_idx, "dyn_not"),
                _ if op == Op::DYN_TO_BOOL => emit_rt_call(&mut body, &rt_idx, "dyn_to_bool"),
                _ if op == Op::STR_CONCAT => emit_rt_call(&mut body, &rt_idx, "str_concat"),

                _ if op == Op::GLOBAL_GET => { let _ = read_u16(&chunk.code, &mut ip); emit_rt_call(&mut body, &rt_idx, "global_get"); }
                _ if op == Op::GLOBAL_SET => { let _ = read_u16(&chunk.code, &mut ip); emit_rt_call(&mut body, &rt_idx, "global_set"); }

                _ if op == Op::STRUCT_GET => { let _ = read_u16(&chunk.code, &mut ip); emit_rt_call(&mut body, &rt_idx, "get_prop"); }
                _ if op == Op::STRUCT_SET => { let _ = read_u16(&chunk.code, &mut ip); emit_rt_call(&mut body, &rt_idx, "set_prop"); }
                _ if op == Op::STRUCT_NEW => { let _ = read_u16(&chunk.code, &mut ip); emit_rt_call(&mut body, &rt_idx, "new_object"); }
                _ if op == Op::ARRAY_GET => emit_rt_call(&mut body, &rt_idx, "array_get"),
                _ if op == Op::ARRAY_SET => emit_rt_call(&mut body, &rt_idx, "array_set"),
                _ if op == Op::ARRAY_NEW => { let _ = read_u16(&chunk.code, &mut ip); emit_rt_call(&mut body, &rt_idx, "new_array"); }

                // --- Skip complex ops, emit nop ---
                _ if op == Op::REF_FUNC => { let _ = read_u16(&chunk.code, &mut ip); let uv = chunk.code[ip] as usize; ip += 1 + uv * 2; body.push(0x01); }
                _ if op == Op::TRY_START => { ip += 4; body.push(0x01); }
                _ if op == Op::UPVALUE_GET || op == Op::UPVALUE_SET => { ip += 1; body.push(0x01); }
                _ if op == Op::STR_CONCAT_N || op == Op::RETURN_CALL || op == Op::CALL_INDIRECT || op == Op::BR_LABEL || op == Op::BR_IF_LABEL => { ip += 1; body.push(0x01); }
                _ if op == Op::BR_TABLE => { let count = chunk.code[ip] as usize; ip += 2 + count; body.push(0x01); }
                _ if op == Op::REF_TEST => { ip += 2; body.push(0x01); }
                _ if op == Op::DUP => body.push(0x01), // nop (can't easily dup in WASM)

                _ => body.push(0x01), // nop
            }
        }
        body.push(0x0B); // end

        write_leb128_u32(&mut out, body.len() as u32);
        out.extend_from_slice(&body);
    }
    out
}

fn emit_rt_call(body: &mut Vec<u8>, rt_idx: &std::collections::HashMap<&str, usize>, name: &str) {
    body.push(0x10); // call
    if let Some(&idx) = rt_idx.get(name) {
        write_leb128_u32(body, idx as u32);
    } else {
        write_leb128_u32(body, 0);
    }
}

// ============================================================
// Reader
// ============================================================

pub fn read_wasm(data: &[u8]) -> Result<Vec<Chunk>, String> {
    if data.len() < 8 || &data[0..4] != &WASM_MAGIC {
        return Err("Invalid WASM: bad magic".into());
    }
    let mut pos = 8;
    let mut custom_data: Option<Vec<u8>> = None;
    let mut type_section: Vec<u8> = Vec::new();
    let mut import_section: Vec<u8> = Vec::new();
    let mut func_section: Vec<u8> = Vec::new();
    let mut export_section: Vec<u8> = Vec::new();
    let mut code_section: Vec<u8> = Vec::new();

    while pos < data.len() {
        if pos >= data.len() { break; }
        let section_id = data[pos]; pos += 1;
        let (size, read) = read_leb128_u32(&data[pos..]);
        pos += read;
        let section_end = (pos + size as usize).min(data.len());
        let section_data = data[pos..section_end].to_vec();

        match section_id {
            SECTION_CUSTOM => {
                // Check if it's our "vybe" custom section
                let (nlen, nr) = read_leb128_u32(&section_data);
                if nlen == 4 && section_data.get(nr..nr+4) == Some(b"vybe") {
                    custom_data = Some(section_data);
                }
            }
            SECTION_TYPE => type_section = section_data,
            SECTION_IMPORT => import_section = section_data,
            SECTION_FUNCTION => func_section = section_data,
            SECTION_EXPORT => export_section = section_data,
            SECTION_CODE => code_section = section_data,
            _ => {} // skip memory, table, global, data, element
        }
        pos = section_end;
    }

    // If we have a vybe custom section, use that for round-trip (our format)
    if let Some(ref cd) = custom_data {
        return decode_vybe_section(cd);
    }

    // Otherwise, decode as standard WASM module
    if code_section.is_empty() {
        return Err("No code section in WASM module".into());
    }
    decode_standard_wasm(&type_section, &import_section, &func_section, &export_section, &code_section)
}

/// Decode a standard WASM module (e.g. from Rust/C compiler)
fn decode_standard_wasm(
    type_sec: &[u8], import_sec: &[u8], _func_sec: &[u8],
    export_sec: &[u8], code_sec: &[u8],
) -> Result<Vec<Chunk>, String> {
    // Parse type section to get function signatures
    let types = parse_type_section(type_sec);

    // Parse imports
    let imports = parse_import_section(import_sec);
    let import_func_count = imports.len();

    // Parse exports to find function names
    let exports = parse_export_section(export_sec);

    // Parse code section
    let mut cpos = 0;
    let (func_count, read) = read_leb128_u32(&code_sec[cpos..]);
    cpos += read;

    let mut chunks = Vec::new();

    // Create a script chunk that calls exported functions
    let mut script = Chunk::new("<script>");
    script.local_count = 0;

    for i in 0..func_count as usize {
        let (body_size, read) = read_leb128_u32(&code_sec[cpos..]);
        cpos += read;
        let _body_start = cpos;
        let body_end = cpos + body_size as usize;

        // Parse locals
        let (local_groups, read) = read_leb128_u32(&code_sec[cpos..]);
        cpos += read;
        let mut local_count: u32 = 0;
        for _ in 0..local_groups {
            let (count, read) = read_leb128_u32(&code_sec[cpos..]);
            cpos += read;
            cpos += 1; // type byte
            local_count += count;
        }

        // Get function name from exports
        let func_idx = import_func_count + i;
        let name = exports.iter()
            .find(|(_, idx)| *idx == func_idx)
            .map(|(n, _)| n.clone())
            .unwrap_or_else(|| format!("func_{}", i));

        // Get arity from type section
        let arity = types.get(func_idx).map(|(params, _)| params.len() as u8).unwrap_or(0);

        // Translate WASM opcodes to our Chunk format
        let wasm_code = &code_sec[cpos..body_end.saturating_sub(1)]; // -1 for trailing 'end'
        let mut chunk = translate_wasm_to_chunk(wasm_code, &name, arity, local_count, import_func_count);
        chunk.emit_op(Op::RETURN, 0);
        chunks.push(chunk);

        cpos = body_end;
    }

    // Add imports to script chunk
    for (module, name, _) in &imports {
        script.add_import(module, name);
    }

    // Insert script as chunk 0
    chunks.insert(0, script);

    Ok(chunks)
}

/// Translate WASM opcodes to our internal Chunk format.
/// Builds a proper constant pool and adjusts local indices.
/// Control flow label for WASM block/loop/if translation
struct WasmLabel {
    /// For blocks: offset of the br placeholder to patch to end
    /// For loops: start offset (br target)
    start_offset: usize,
    is_loop: bool,
    /// Pending forward jumps to patch when we hit 'end'
    break_patches: Vec<usize>,
    /// For if/else: the br_if_false patch location
    if_patch: Option<usize>,
}

fn translate_wasm_to_chunk(wasm: &[u8], name: &str, arity: u8, wasm_local_count: u32, _import_count: usize) -> Chunk {
    let mut chunk = Chunk::new(name);
    chunk.arity = arity;
    chunk.local_count = 1 + arity as u16 + wasm_local_count as u16;

    let mut pos = 0;
    let mut label_stack: Vec<WasmLabel> = Vec::new();

    while pos < wasm.len() {
        let byte = wasm[pos]; pos += 1;

        match byte {
            0x00 => chunk.emit_op(Op::HALT, 0),
            0x01 => {} // nop

            // block blocktype — forward jump target
            0x02 => {
                skip_leb128(wasm, &mut pos); // blocktype
                label_stack.push(WasmLabel {
                    start_offset: chunk.current_offset(),
                    is_loop: false,
                    break_patches: Vec::new(),
                    if_patch: None,
                });
            }

            // loop blocktype — backward jump target
            0x03 => {
                skip_leb128(wasm, &mut pos);
                label_stack.push(WasmLabel {
                    start_offset: chunk.current_offset(),
                    is_loop: true,
                    break_patches: Vec::new(),
                    if_patch: None,
                });
            }

            // if blocktype — conditional block
            0x04 => {
                skip_leb128(wasm, &mut pos);
                chunk.emit_op(Op::DYN_TO_BOOL, 0);
                let patch = chunk.emit_jump(Op::BR_IF_FALSE, 0);
                label_stack.push(WasmLabel {
                    start_offset: chunk.current_offset(),
                    is_loop: false,
                    break_patches: Vec::new(),
                    if_patch: Some(patch),
                });
            }

            // else
            0x05 => {
                if let Some(label) = label_stack.last_mut() {
                    // Jump over else block from end of if-true
                    let skip = chunk.emit_jump(Op::BR, 0);
                    label.break_patches.push(skip);
                    // Patch the if_patch to here (start of else)
                    if let Some(patch) = label.if_patch.take() {
                        chunk.patch_jump(patch);
                    }
                }
            }

            // end
            0x0B => {
                if let Some(label) = label_stack.pop() {
                    let _end_offset = chunk.current_offset();
                    // Patch if_patch if not already patched (no else branch)
                    if let Some(patch) = label.if_patch {
                        chunk.patch_jump(patch);
                    }
                    // Patch all forward break jumps to here
                    for patch in &label.break_patches {
                        chunk.patch_jump(*patch);
                    }
                }
            }

            // br N — branch to Nth enclosing label
            0x0C => {
                let (depth, _) = read_leb128_u32(&wasm[pos..]);
                skip_leb128(wasm, &mut pos);
                let depth = depth as usize;
                if let Some(label) = label_stack.iter().rev().nth(depth) {
                    if label.is_loop {
                        // Loop: jump back to start
                        chunk.emit_loop(label.start_offset, 0);
                    } else {
                        // Block: jump forward to end (patch later)
                        let patch = chunk.emit_jump(Op::BR, 0);
                        // Store patch in the target label
                        let idx = label_stack.len() - 1 - depth;
                        label_stack[idx].break_patches.push(patch);
                    }
                }
            }

            // br_if N — conditional branch
            0x0D => {
                let (depth, _) = read_leb128_u32(&wasm[pos..]);
                skip_leb128(wasm, &mut pos);
                let depth = depth as usize;
                chunk.emit_op(Op::DYN_TO_BOOL, 0);
                if let Some(label) = label_stack.iter().rev().nth(depth) {
                    if label.is_loop {
                        // Loop: conditional jump back
                        let exit = chunk.emit_jump(Op::BR_IF_FALSE, 0);
                        chunk.emit_loop(label.start_offset, 0);
                        chunk.patch_jump(exit);
                    } else {
                        // Block: conditional jump forward
                        let patch = chunk.emit_jump(Op::BR_IF_TRUE, 0);
                        let idx = label_stack.len() - 1 - depth;
                        label_stack[idx].break_patches.push(patch);
                    }
                }
            }
            0x0E => {
                // br_table — branch table
                let (count, _) = read_leb128_u32(&wasm[pos..]);
                skip_leb128(wasm, &mut pos);
                for _ in 0..=count { skip_leb128(wasm, &mut pos); } // skip all labels + default
                chunk.emit_op(Op::DROP, 0); // simplified
            }
            0x0F => chunk.emit_op(Op::RETURN, 0),
            0x1A => chunk.emit_op(Op::DROP, 0),
            0x1B => chunk.emit_op(Op::SELECT, 0),

            // call — adjust index (skip imports, offset to our chunk indices)
            0x10 => {
                let (idx, _) = read_leb128_u32(&wasm[pos..]);
                skip_leb128(wasm, &mut pos);
                chunk.emit_op_u8(Op::CALL, idx as u8, 0);
            }

            // local.get — offset by 1 (our slot 0 = implicit fn)
            0x20 => {
                let (idx, _) = read_leb128_u32(&wasm[pos..]);
                skip_leb128(wasm, &mut pos);
                chunk.emit_op_u16(Op::LOCAL_GET, (idx as u16) + 1, 0);
            }
            // local.set
            0x21 => {
                let (idx, _) = read_leb128_u32(&wasm[pos..]);
                skip_leb128(wasm, &mut pos);
                chunk.emit_op_u16(Op::LOCAL_SET, (idx as u16) + 1, 0);
            }
            // local.tee
            0x22 => {
                let (idx, _) = read_leb128_u32(&wasm[pos..]);
                skip_leb128(wasm, &mut pos);
                chunk.emit_op_u16(Op::LOCAL_SET, (idx as u16) + 1, 0);
            }

            // i32.const → add to constant pool as F64
            0x41 => {
                let (val, read) = read_leb128_i32(&wasm[pos..]);
                pos += read;
                let ci = chunk.add_constant(Value::F64(val as f64));
                chunk.emit_op_u16(Op::CONST, ci, 0);
            }

            // i64.const
            0x42 => {
                let (val, read) = read_leb128_i64(&wasm[pos..]);
                pos += read;
                let ci = chunk.add_constant(Value::F64(val as f64));
                chunk.emit_op_u16(Op::CONST, ci, 0);
            }

            // f64.const
            0x44 => {
                if pos + 8 <= wasm.len() {
                    let val = f64::from_le_bytes([
                        wasm[pos], wasm[pos+1], wasm[pos+2], wasm[pos+3],
                        wasm[pos+4], wasm[pos+5], wasm[pos+6], wasm[pos+7],
                    ]);
                    pos += 8;
                    let ci = chunk.add_constant(Value::F64(val));
                    chunk.emit_op_u16(Op::CONST, ci, 0);
                }
            }

            // f32.const
            0x43 => {
                if pos + 4 <= wasm.len() {
                    let val = f32::from_le_bytes([wasm[pos], wasm[pos+1], wasm[pos+2], wasm[pos+3]]);
                    pos += 4;
                    let ci = chunk.add_constant(Value::F64(val as f64));
                    chunk.emit_op_u16(Op::CONST, ci, 0);
                }
            }

            // i32 arithmetic — ALL opcodes
            0x67 => chunk.emit_op(Op::I32_CLZ, 0),
            0x68 => chunk.emit_op(Op::I32_CTZ, 0),
            0x69 => chunk.emit_op(Op::I32_POPCNT, 0),
            0x6A => chunk.emit_op(Op::I32_ADD, 0),
            0x6B => chunk.emit_op(Op::I32_SUB, 0),
            0x6C => chunk.emit_op(Op::I32_MUL, 0),
            0x6D => chunk.emit_op(Op::I32_DIV_S, 0),
            0x6E => chunk.emit_op(Op::I32_DIV_U, 0),
            0x6F => chunk.emit_op(Op::I32_REM_S, 0),
            0x70 => chunk.emit_op(Op::I32_REM_U, 0),
            0x71 => chunk.emit_op(Op::I32_AND, 0),
            0x72 => chunk.emit_op(Op::I32_OR, 0),
            0x73 => chunk.emit_op(Op::I32_XOR, 0),
            0x74 => chunk.emit_op(Op::I32_SHL, 0),
            0x75 => chunk.emit_op(Op::I32_SHR_S, 0),
            0x76 => chunk.emit_op(Op::I32_SHR_U, 0),
            0x77 => chunk.emit_op(Op::I32_ROTL, 0),
            0x78 => chunk.emit_op(Op::I32_ROTR, 0),

            // i64 arithmetic — ALL opcodes
            0x79 => chunk.emit_op(Op::I64_CLZ, 0),
            0x7A => chunk.emit_op(Op::I64_CTZ, 0),
            0x7B => chunk.emit_op(Op::I64_POPCNT, 0),
            0x7C => chunk.emit_op(Op::I64_ADD, 0),
            0x7D => chunk.emit_op(Op::I64_SUB, 0),
            0x7E => chunk.emit_op(Op::I64_MUL, 0),
            0x7F => chunk.emit_op(Op::I64_DIV_S, 0),
            0x80 => chunk.emit_op(Op::I64_DIV_U, 0),
            0x81 => chunk.emit_op(Op::I64_REM_S, 0),
            0x82 => chunk.emit_op(Op::I64_REM_U, 0),
            0x83 => chunk.emit_op(Op::I64_AND, 0),
            0x84 => chunk.emit_op(Op::I64_OR, 0),
            0x85 => chunk.emit_op(Op::I64_XOR, 0),
            0x86 => chunk.emit_op(Op::I64_SHL, 0),
            0x87 => chunk.emit_op(Op::I64_SHR_S, 0),
            0x88 => chunk.emit_op(Op::I64_SHR_U, 0),
            0x89 => chunk.emit_op(Op::I64_ROTL, 0),
            0x8A => chunk.emit_op(Op::I64_ROTR, 0),

            // i64 comparison
            0x50 => chunk.emit_op(Op::I64_EQZ, 0),
            0x51 => chunk.emit_op(Op::DYN_EQ, 0),    // i64.eq
            0x52 => chunk.emit_op(Op::DYN_NE, 0),    // i64.ne
            0x53 => chunk.emit_op(Op::DYN_LT, 0),    // i64.lt_s
            0x54 => chunk.emit_op(Op::DYN_LT, 0),    // i64.lt_u
            0x55 => chunk.emit_op(Op::DYN_GT, 0),    // i64.gt_s
            0x56 => chunk.emit_op(Op::DYN_GT, 0),    // i64.gt_u
            0x57 => chunk.emit_op(Op::DYN_LE, 0),    // i64.le_s
            0x58 => chunk.emit_op(Op::DYN_LE, 0),    // i64.le_u
            0x59 => chunk.emit_op(Op::DYN_GE, 0),    // i64.ge_s
            0x5A => chunk.emit_op(Op::DYN_GE, 0),    // i64.ge_u

            // i32 comparison
            0x45 => chunk.emit_op(Op::I32_EQZ, 0),
            0x46 => chunk.emit_op(Op::DYN_EQ, 0),     // i32.eq
            0x47 => chunk.emit_op(Op::DYN_NE, 0),     // i32.ne
            0x48 => chunk.emit_op(Op::DYN_LT, 0),     // i32.lt_s
            0x49 => chunk.emit_op(Op::DYN_LT, 0),     // i32.lt_u
            0x4A => chunk.emit_op(Op::DYN_GT, 0),     // i32.gt_s
            0x4B => chunk.emit_op(Op::DYN_GT, 0),     // i32.gt_u
            0x4C => chunk.emit_op(Op::DYN_LE, 0),     // i32.le_s
            0x4D => chunk.emit_op(Op::DYN_LE, 0),     // i32.le_u
            0x4E => chunk.emit_op(Op::DYN_GE, 0),     // i32.ge_s
            0x4F => chunk.emit_op(Op::DYN_GE, 0),     // i32.ge_u

            // f64 arithmetic — ALL opcodes
            0xA0 => chunk.emit_op(Op::F64_ADD, 0),
            0xA1 => chunk.emit_op(Op::F64_SUB, 0),
            0xA2 => chunk.emit_op(Op::F64_MUL, 0),
            0xA3 => chunk.emit_op(Op::F64_DIV, 0),
            0xA4 => chunk.emit_op(Op::F64_MIN, 0),
            0xA5 => chunk.emit_op(Op::F64_MAX, 0),
            0xA6 => chunk.emit_op(Op::F64_COPYSIGN, 0),

            // f32 comparison
            0x5B => chunk.emit_op(Op::DYN_EQ, 0),    // f32.eq
            0x5C => chunk.emit_op(Op::DYN_NE, 0),    // f32.ne
            0x5D => chunk.emit_op(Op::DYN_LT, 0),    // f32.lt
            0x5E => chunk.emit_op(Op::DYN_GT, 0),    // f32.gt
            0x5F => chunk.emit_op(Op::DYN_LE, 0),    // f32.le
            0x60 => chunk.emit_op(Op::DYN_GE, 0),    // f32.ge

            // f64 comparison
            0x61 => chunk.emit_op(Op::DYN_EQ, 0),    // f64.eq
            0x62 => chunk.emit_op(Op::DYN_NE, 0),    // f64.ne
            0x63 => chunk.emit_op(Op::DYN_LT, 0),    // f64.lt
            0x64 => chunk.emit_op(Op::DYN_GT, 0),    // f64.gt
            0x65 => chunk.emit_op(Op::DYN_LE, 0),    // f64.le
            0x66 => chunk.emit_op(Op::DYN_GE, 0),    // f64.ge

            // Memory — ALL load/store opcodes
            0x28 => { skip_leb128(wasm, &mut pos); skip_leb128(wasm, &mut pos); chunk.emit_op(Op::I32_LOAD, 0); }
            0x29 => { skip_leb128(wasm, &mut pos); skip_leb128(wasm, &mut pos); chunk.emit_op(Op::I64_LOAD, 0); }
            0x2A => { skip_leb128(wasm, &mut pos); skip_leb128(wasm, &mut pos); chunk.emit_op(Op::F32_LOAD, 0); }
            0x2B => { skip_leb128(wasm, &mut pos); skip_leb128(wasm, &mut pos); chunk.emit_op(Op::F64_LOAD, 0); }
            0x2C => { skip_leb128(wasm, &mut pos); skip_leb128(wasm, &mut pos); chunk.emit_op(Op::I32_LOAD8_S, 0); }
            0x2D => { skip_leb128(wasm, &mut pos); skip_leb128(wasm, &mut pos); chunk.emit_op(Op::I32_LOAD8_U, 0); }
            0x2E => { skip_leb128(wasm, &mut pos); skip_leb128(wasm, &mut pos); chunk.emit_op(Op::I32_LOAD16_S, 0); }
            0x2F => { skip_leb128(wasm, &mut pos); skip_leb128(wasm, &mut pos); chunk.emit_op(Op::I32_LOAD16_U, 0); }
            0x30 => { skip_leb128(wasm, &mut pos); skip_leb128(wasm, &mut pos); chunk.emit_op(Op::I64_LOAD8_S, 0); }
            0x31 => { skip_leb128(wasm, &mut pos); skip_leb128(wasm, &mut pos); chunk.emit_op(Op::I64_LOAD8_U, 0); }
            0x32 => { skip_leb128(wasm, &mut pos); skip_leb128(wasm, &mut pos); chunk.emit_op(Op::I64_LOAD16_S, 0); }
            0x33 => { skip_leb128(wasm, &mut pos); skip_leb128(wasm, &mut pos); chunk.emit_op(Op::I64_LOAD16_U, 0); }
            0x34 => { skip_leb128(wasm, &mut pos); skip_leb128(wasm, &mut pos); chunk.emit_op(Op::I64_LOAD32_S, 0); }
            0x35 => { skip_leb128(wasm, &mut pos); skip_leb128(wasm, &mut pos); chunk.emit_op(Op::I64_LOAD32_U, 0); }
            0x36 => { skip_leb128(wasm, &mut pos); skip_leb128(wasm, &mut pos); chunk.emit_op(Op::I32_STORE, 0); }
            0x37 => { skip_leb128(wasm, &mut pos); skip_leb128(wasm, &mut pos); chunk.emit_op(Op::I64_STORE, 0); }
            0x38 => { skip_leb128(wasm, &mut pos); skip_leb128(wasm, &mut pos); chunk.emit_op(Op::F32_STORE, 0); }
            0x39 => { skip_leb128(wasm, &mut pos); skip_leb128(wasm, &mut pos); chunk.emit_op(Op::F64_STORE, 0); }
            0x3A => { skip_leb128(wasm, &mut pos); skip_leb128(wasm, &mut pos); chunk.emit_op(Op::I32_STORE8, 0); }
            0x3B => { skip_leb128(wasm, &mut pos); skip_leb128(wasm, &mut pos); chunk.emit_op(Op::I32_STORE16, 0); }
            0x3C => { skip_leb128(wasm, &mut pos); skip_leb128(wasm, &mut pos); chunk.emit_op(Op::I64_STORE8, 0); }
            0x3D => { skip_leb128(wasm, &mut pos); skip_leb128(wasm, &mut pos); chunk.emit_op(Op::I64_STORE16, 0); }
            0x3E => { skip_leb128(wasm, &mut pos); skip_leb128(wasm, &mut pos); chunk.emit_op(Op::I64_STORE32, 0); }
            0x3F => { skip_leb128(wasm, &mut pos); chunk.emit_op(Op::MEMORY_SIZE, 0); }
            0x40 => { skip_leb128(wasm, &mut pos); chunk.emit_op(Op::MEMORY_GROW, 0); }

            // f32 arithmetic — ALL opcodes
            0x8B => chunk.emit_op(Op::F32_ABS, 0),
            0x8C => chunk.emit_op(Op::F32_NEG, 0),
            0x8D => chunk.emit_op(Op::F32_CEIL, 0),
            0x8E => chunk.emit_op(Op::F32_FLOOR, 0),
            0x8F => chunk.emit_op(Op::F32_TRUNC, 0),
            0x90 => chunk.emit_op(Op::F32_NEAREST, 0),
            0x91 => chunk.emit_op(Op::F32_SQRT, 0),
            0x92 => chunk.emit_op(Op::F64_ADD, 0),   // f32.add (promoted)
            0x93 => chunk.emit_op(Op::F64_SUB, 0),   // f32.sub (promoted)
            0x94 => chunk.emit_op(Op::F64_MUL, 0),   // f32.mul (promoted)
            0x95 => chunk.emit_op(Op::F64_DIV, 0),   // f32.div (promoted)
            0x96 => chunk.emit_op(Op::F32_MIN, 0),
            0x97 => chunk.emit_op(Op::F32_MAX, 0),
            0x98 => chunk.emit_op(Op::F32_COPYSIGN, 0),

            // f64 extra ops — ALL opcodes
            0x99 => chunk.emit_op(Op::F64_ABS, 0),
            0x9A => chunk.emit_op(Op::F64_NEG, 0),
            0x9B => chunk.emit_op(Op::F64_CEIL, 0),
            0x9C => chunk.emit_op(Op::F64_FLOOR, 0),
            0x9D => chunk.emit_op(Op::F64_TRUNC, 0),
            0x9E => chunk.emit_op(Op::F64_NEAREST, 0),
            0x9F => chunk.emit_op(Op::F64_SQRT, 0),

            // ALL conversions
            0xA7 => chunk.emit_op(Op::I32_WRAP_I64, 0),
            0xA8 => chunk.emit_op(Op::I32_FROM_F64, 0), // i32.trunc_f32_s
            0xA9 => chunk.emit_op(Op::I32_FROM_F64, 0), // i32.trunc_f32_u
            0xAA => chunk.emit_op(Op::I32_FROM_F64, 0), // i32.trunc_f64_s
            0xAB => chunk.emit_op(Op::I32_FROM_F64, 0), // i32.trunc_f64_u
            0xAC => chunk.emit_op(Op::I64_EXTEND_I32_S, 0),
            0xAD => chunk.emit_op(Op::I64_EXTEND_I32_U, 0),
            0xAE => chunk.emit_op(Op::I64_TRUNC_F64_S, 0), // i64.trunc_f32_s (f32=f64 in VM)
            0xAF => chunk.emit_op(Op::I64_TRUNC_F64_U, 0), // i64.trunc_f32_u
            0xB0 => chunk.emit_op(Op::I64_TRUNC_F64_S, 0),
            0xB1 => chunk.emit_op(Op::I64_TRUNC_F64_U, 0),
            0xB2 => chunk.emit_op(Op::F64_FROM_I32, 0), // f32.convert_i32_s
            0xB3 => chunk.emit_op(Op::F64_FROM_I32, 0), // f32.convert_i32_u
            0xB4 => chunk.emit_op(Op::F64_FROM_I32, 0), // f32.convert_i64_s (i64→f64)
            0xB5 => chunk.emit_op(Op::F64_FROM_I32, 0), // f32.convert_i64_u
            0xB6 => chunk.emit_op(Op::F32_DEMOTE_F64, 0),
            0xB7 => chunk.emit_op(Op::F64_FROM_I32, 0), // f64.convert_i32_s
            0xB8 => chunk.emit_op(Op::F64_FROM_I32, 0), // f64.convert_i32_u
            0xB9 => chunk.emit_op(Op::F64_PROMOTE_F32, 0),
            0xBA => chunk.emit_op(Op::F32_REINTERPRET_I32, 0),
            0xBB => chunk.emit_op(Op::F64_REINTERPRET_I64, 0),
            0xBC => chunk.emit_op(Op::I32_REINTERPRET_F32, 0),
            0xBD => chunk.emit_op(Op::I64_REINTERPRET_F64, 0),

            // Sign extension
            0xC0 => chunk.emit_op(Op::I32_EXTEND8_S, 0),
            0xC1 => chunk.emit_op(Op::I32_EXTEND16_S, 0),
            0xC2 => chunk.emit_op(Op::I64_EXTEND8_S, 0),
            0xC3 => chunk.emit_op(Op::I64_EXTEND16_S, 0),
            0xC4 => chunk.emit_op(Op::I64_EXTEND32_S, 0),

            // global.get/set — WASM globals mapped to global_get/set with index as name
            0x23 => {
                let (idx, _) = read_leb128_u32(&wasm[pos..]);
                skip_leb128(wasm, &mut pos);
                let name = format!("__wasm_global_{}", idx);
                let ci = chunk.add_constant(Value::String(Arc::from(name.as_str())));
                chunk.emit_op_u16(Op::GLOBAL_GET, ci, 0);
            }
            0x24 => {
                let (idx, _) = read_leb128_u32(&wasm[pos..]);
                skip_leb128(wasm, &mut pos);
                let name = format!("__wasm_global_{}", idx);
                let ci = chunk.add_constant(Value::String(Arc::from(name.as_str())));
                chunk.emit_op_u16(Op::GLOBAL_SET, ci, 0);
            }

            // call_indirect
            0x11 => { skip_leb128(wasm, &mut pos); skip_leb128(wasm, &mut pos); }

            // Unknown — skip
            _ => {}
        }
    }

    chunk
}

fn read_leb128_i64(data: &[u8]) -> (i64, usize) {
    let mut result = 0i64;
    let mut shift = 0;
    let mut pos = 0;
    loop {
        if pos >= data.len() { break; }
        let byte = data[pos]; pos += 1;
        result |= ((byte & 0x7F) as i64) << shift;
        shift += 7;
        if byte & 0x80 == 0 {
            if shift < 64 && (byte & 0x40 != 0) {
                result |= !0i64 << shift;
            }
            break;
        }
    }
    (result, pos)
}

fn parse_type_section(data: &[u8]) -> Vec<(Vec<u8>, Vec<u8>)> {
    if data.is_empty() { return vec![]; }
    let mut pos = 0;
    let (count, read) = read_leb128_u32(&data[pos..]);
    pos += read;
    let mut types = Vec::new();
    for _ in 0..count {
        if pos >= data.len() || data[pos] != TYPE_FUNC { pos += 1; continue; }
        pos += 1; // skip 0x60
        let (param_count, read) = read_leb128_u32(&data[pos..]);
        pos += read;
        let params: Vec<u8> = data[pos..pos + param_count as usize].to_vec();
        pos += param_count as usize;
        let (result_count, read) = read_leb128_u32(&data[pos..]);
        pos += read;
        let results: Vec<u8> = data[pos..pos + result_count as usize].to_vec();
        pos += result_count as usize;
        types.push((params, results));
    }
    types
}

fn parse_import_section(data: &[u8]) -> Vec<(String, String, u8)> {
    if data.is_empty() { return vec![]; }
    let mut pos = 0;
    let (count, read) = read_leb128_u32(&data[pos..]);
    pos += read;
    let mut imports = Vec::new();
    for _ in 0..count {
        let (mlen, read) = read_leb128_u32(&data[pos..]); pos += read;
        let module = std::str::from_utf8(&data[pos..pos + mlen as usize]).unwrap_or("").to_string();
        pos += mlen as usize;
        let (nlen, read) = read_leb128_u32(&data[pos..]); pos += read;
        let name = std::str::from_utf8(&data[pos..pos + nlen as usize]).unwrap_or("").to_string();
        pos += nlen as usize;
        let kind = data[pos]; pos += 1;
        skip_leb128(&data, &mut pos); // type index or other descriptor
        imports.push((module, name, kind));
    }
    imports
}

fn parse_export_section(data: &[u8]) -> Vec<(String, usize)> {
    if data.is_empty() { return vec![]; }
    let mut pos = 0;
    let (count, read) = read_leb128_u32(&data[pos..]);
    pos += read;
    let mut exports = Vec::new();
    for _ in 0..count {
        let (nlen, read) = read_leb128_u32(&data[pos..]); pos += read;
        let name = std::str::from_utf8(&data[pos..pos + nlen as usize]).unwrap_or("").to_string();
        pos += nlen as usize;
        let kind = data[pos]; pos += 1;
        let (idx, read) = read_leb128_u32(&data[pos..]); pos += read;
        if kind == 0 { // function export
            exports.push((name, idx as usize));
        }
    }
    exports
}

fn skip_leb128(data: &[u8], pos: &mut usize) {
    while *pos < data.len() {
        let byte = data[*pos]; *pos += 1;
        if byte & 0x80 == 0 { break; }
    }
}

fn read_leb128_i32(data: &[u8]) -> (i32, usize) {
    let mut result = 0i32;
    let mut shift = 0;
    let mut pos = 0;
    loop {
        if pos >= data.len() { break; }
        let byte = data[pos]; pos += 1;
        result |= ((byte & 0x7F) as i32) << shift;
        shift += 7;
        if byte & 0x80 == 0 {
            if shift < 32 && (byte & 0x40 != 0) {
                result |= !0 << shift;
            }
            break;
        }
    }
    (result, pos)
}

fn decode_vybe_section(data: &[u8]) -> Result<Vec<Chunk>, String> {
    let mut pos = 0;
    let (name_len, read) = read_leb128_u32(&data[pos..]);
    pos += read;
    let _name = &data[pos..pos + name_len as usize];
    pos += name_len as usize;

    let (count, read) = read_leb128_u32(&data[pos..]);
    pos += read;

    let mut chunks = Vec::new();
    for _ in 0..count {
        let (nlen, read) = read_leb128_u32(&data[pos..]); pos += read;
        let name = std::str::from_utf8(&data[pos..pos + nlen as usize]).unwrap_or("").to_string();
        pos += nlen as usize;
        let arity = data[pos]; pos += 1;
        let (lc, read) = read_leb128_u32(&data[pos..]); pos += read;

        // Constants
        let (cc, read) = read_leb128_u32(&data[pos..]); pos += read;
        let mut constants = Vec::new();
        for _ in 0..cc { constants.push(decode_value(data, &mut pos)); }

        // Imports
        let (ic, read) = read_leb128_u32(&data[pos..]); pos += read;
        let mut imports = Vec::new();
        for _ in 0..ic {
            let (mlen, read) = read_leb128_u32(&data[pos..]); pos += read;
            let module = std::str::from_utf8(&data[pos..pos + mlen as usize]).unwrap_or("").to_string();
            pos += mlen as usize;
            let (nlen, read) = read_leb128_u32(&data[pos..]); pos += read;
            let iname = std::str::from_utf8(&data[pos..pos + nlen as usize]).unwrap_or("").to_string();
            pos += nlen as usize;
            imports.push(crate::chunk::Import { module, name: iname });
        }

        // Bytecode
        let (code_len, read) = read_leb128_u32(&data[pos..]); pos += read;
        let code = data[pos..pos + code_len as usize].to_vec();
        pos += code_len as usize;

        let mut chunk = Chunk::new(&name);
        chunk.arity = arity;
        chunk.local_count = lc as u16;
        chunk.constants = constants;
        chunk.imports = imports;
        chunk.code = code;
        chunks.push(chunk);
    }
    Ok(chunks)
}

// ============================================================
// Custom section: our metadata + raw bytecode for round-trip
// ============================================================

fn encode_custom_section(chunks: &[Chunk]) -> Vec<u8> {
    let mut out = Vec::new();
    write_name(&mut out, "vybe");
    write_leb128_u32(&mut out, chunks.len() as u32);
    for chunk in chunks {
        write_name(&mut out, &chunk.name);
        out.push(chunk.arity);
        write_leb128_u32(&mut out, chunk.local_count as u32);
        write_leb128_u32(&mut out, chunk.constants.len() as u32);
        for c in &chunk.constants { encode_value(&mut out, c); }
        write_leb128_u32(&mut out, chunk.imports.len() as u32);
        for imp in &chunk.imports {
            write_name(&mut out, &imp.module);
            write_name(&mut out, &imp.name);
        }
        // Store raw bytecode for round-trip
        write_leb128_u32(&mut out, chunk.code.len() as u32);
        out.extend_from_slice(&chunk.code);
    }
    out
}

// ============================================================
// Helpers
// ============================================================

fn write_section(out: &mut Vec<u8>, id: u8, data: &[u8]) {
    out.push(id);
    write_leb128_u32(out, data.len() as u32);
    out.extend_from_slice(data);
}

fn write_name(out: &mut Vec<u8>, s: &str) {
    write_leb128_u32(out, s.len() as u32);
    out.extend_from_slice(s.as_bytes());
}

fn write_leb128_u32(out: &mut Vec<u8>, mut value: u32) {
    loop {
        let mut byte = (value & 0x7F) as u8;
        value >>= 7;
        if value != 0 { byte |= 0x80; }
        out.push(byte);
        if value == 0 { break; }
    }
}

fn write_leb128_i32(out: &mut Vec<u8>, mut value: i32) {
    loop {
        let mut byte = (value & 0x7F) as u8;
        value >>= 7;
        let more = !(((value == 0) && (byte & 0x40 == 0)) || ((value == -1) && (byte & 0x40 != 0)));
        if more { byte |= 0x80; }
        out.push(byte);
        if !more { break; }
    }
}

fn write_leb128_i64(out: &mut Vec<u8>, mut value: i64) {
    loop {
        let mut byte = (value & 0x7F) as u8;
        value >>= 7;
        let more = !(((value == 0) && (byte & 0x40 == 0)) || ((value == -1) && (byte & 0x40 != 0)));
        if more { byte |= 0x80; }
        out.push(byte);
        if !more { break; }
    }
}

fn read_leb128_u32(data: &[u8]) -> (u32, usize) {
    let mut result = 0u32;
    let mut shift = 0;
    let mut pos = 0;
    loop {
        if pos >= data.len() { break; }
        let byte = data[pos]; pos += 1;
        result |= ((byte & 0x7F) as u32) << shift;
        shift += 7;
        if byte & 0x80 == 0 { break; }
    }
    (result, pos)
}

fn read_u16(code: &[u8], ip: &mut usize) -> u16 {
    let hi = code.get(*ip).copied().unwrap_or(0) as u16;
    let lo = code.get(*ip + 1).copied().unwrap_or(0) as u16;
    *ip += 2;
    (hi << 8) | lo
}

fn read_i16(code: &[u8], ip: &mut usize) -> i16 {
    read_u16(code, ip) as i16
}

fn encode_value(out: &mut Vec<u8>, val: &Value) {
    match val {
        Value::Null | Value::Undefined => out.push(0),
        Value::Bool(b) => { out.push(1); out.push(if *b { 1 } else { 0 }); }
        Value::I32(n) => { out.push(2); out.extend_from_slice(&n.to_le_bytes()); }
        Value::I64(n) => { out.push(3); out.extend_from_slice(&n.to_le_bytes()); }
        Value::F64(n) => { out.push(4); out.extend_from_slice(&n.to_le_bytes()); }
        Value::String(s) => { out.push(5); write_name(out, s); }
        Value::Object(_) | Value::WeakRef(_) => out.push(0),
        Value::V128(b) => { out.push(6); out.extend_from_slice(b); }
    }
}

fn decode_value(data: &[u8], pos: &mut usize) -> Value {
    if *pos >= data.len() { return Value::Null; }
    let tag = data[*pos]; *pos += 1;
    match tag {
        1 => { let b = data[*pos]; *pos += 1; Value::Bool(b != 0) }
        2 => { let n = i32::from_le_bytes([data[*pos],data[*pos+1],data[*pos+2],data[*pos+3]]); *pos += 4; Value::I32(n) }
        3 => { let n = i64::from_le_bytes([data[*pos],data[*pos+1],data[*pos+2],data[*pos+3],data[*pos+4],data[*pos+5],data[*pos+6],data[*pos+7]]); *pos += 8; Value::I64(n) }
        4 => { let n = f64::from_le_bytes([data[*pos],data[*pos+1],data[*pos+2],data[*pos+3],data[*pos+4],data[*pos+5],data[*pos+6],data[*pos+7]]); *pos += 8; Value::F64(n) }
        5 => { let (len, read) = read_leb128_u32(&data[*pos..]); *pos += read; let s = std::str::from_utf8(&data[*pos..*pos+len as usize]).unwrap_or(""); *pos += len as usize; Value::String(Arc::from(s)) }
        _ => Value::Null,
    }
}

/// Calculate total bytes an opcode + operands consume in our internal format
/// Total instruction size: 2-byte opcode + operand bytes.
fn opcode_size(op: Op, code: &[u8], ip: usize) -> usize {
    use crate::opcode::OperandFormat;
    let base = 2; // all opcodes are 2 bytes
    match op.operand_format() {
        OperandFormat::Closure => {
            // ref_func: 2 op + 2 func_idx + 1 uv_count + uv_count*2 descriptors
            let uv_count = code.get(ip + 4).copied().unwrap_or(0) as usize;
            base + 2 + 1 + uv_count * 2
        }
        OperandFormat::BrTable => {
            // br_table: 2 op + 1 count + 1 default + count labels
            let count = code.get(ip + 2).copied().unwrap_or(0) as usize;
            base + 2 + count
        }
        OperandFormat::TryTable => {
            // try_table: 2 op + 1 count + count * 3
            let count = code.get(ip + 2).copied().unwrap_or(0) as usize;
            base + 1 + count * 3
        }
        fmt => base + fmt.fixed_size(),
    }
}
