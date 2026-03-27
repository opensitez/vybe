//! WASM binary format reader/writer.
//!
//! Produces valid WASM modules that tools like wasm-dis can parse.
//! Dynamic operations (dyn_add, struct_get, globals) go through
//! imported vybe:rt functions. Standard WASM ops use real opcodes.

use crate::{Chunk, Op};
use crate::value::Value;
use std::rc::Rc;

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
const TYPE_I32: u8 = 0x7F;
const TYPE_I64: u8 = 0x7E;
const TYPE_F64: u8 = 0x7C;
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
            if let Some(op) = Op::from_byte(chunk.code[ip]) {
                match op {
                    Op::dyn_add => { needed.insert("dyn_add"); }
                    Op::dyn_eq => { needed.insert("dyn_eq"); }
                    Op::dyn_ne => { needed.insert("dyn_ne"); }
                    Op::dyn_lt => { needed.insert("dyn_lt"); }
                    Op::dyn_gt => { needed.insert("dyn_gt"); }
                    Op::dyn_le => { needed.insert("dyn_le"); }
                    Op::dyn_ge => { needed.insert("dyn_ge"); }
                    Op::dyn_neg => { needed.insert("dyn_neg"); }
                    Op::dyn_not => { needed.insert("dyn_not"); }
                    Op::dyn_to_bool => { needed.insert("dyn_to_bool"); }
                    Op::str_concat => { needed.insert("str_concat"); }
                    Op::struct_get => { needed.insert("get_prop"); }
                    Op::struct_set => { needed.insert("set_prop"); }
                    Op::struct_new => { needed.insert("new_object"); }
                    Op::array_get => { needed.insert("array_get"); }
                    Op::array_set => { needed.insert("array_set"); }
                    Op::array_new => { needed.insert("new_array"); }
                    Op::global_get => { needed.insert("global_get"); }
                    Op::global_set => { needed.insert("global_set"); }
                    _ => {}
                }
                ip += opcode_size(op, &chunk.code, ip);
            } else {
                ip += 1;
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

fn encode_export_section(chunks: &[Chunk], import_count: usize) -> Vec<u8> {
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
            let op = match Op::from_byte(chunk.code[ip]) {
                Some(op) => op,
                None => { ip += 1; continue; }
            };
            ip += 1;

            match op {
                // --- Standard WASM ops ---
                Op::local_get => { body.push(0x20); write_leb128_u32(&mut body, read_u16(&chunk.code, &mut ip) as u32); }
                Op::local_set => { body.push(0x22); write_leb128_u32(&mut body, read_u16(&chunk.code, &mut ip) as u32); } // tee
                Op::drop => body.push(0x1A),
                Op::r#return => body.push(0x0F),
                Op::halt => body.push(0x00), // unreachable

                Op::f64_add => body.push(0xA0),
                Op::f64_sub => body.push(0xA1),
                Op::f64_mul => body.push(0xA2),
                Op::f64_div => body.push(0xA3),
                Op::i32_add => body.push(0x6A),
                Op::i32_sub => body.push(0x6B),
                Op::i32_mul => body.push(0x6C),
                Op::i32_and => body.push(0x71),
                Op::i32_or => body.push(0x72),
                Op::i32_xor => body.push(0x73),
                Op::i32_shl => body.push(0x74),
                Op::i32_shr_s => body.push(0x75),

                Op::memory_size => { body.push(0x3F); body.push(0x00); }
                Op::memory_grow => { body.push(0x40); body.push(0x00); }
                Op::i32_load => { body.push(0x28); body.push(0x02); body.push(0x00); }
                Op::i32_store => { body.push(0x36); body.push(0x02); body.push(0x00); }
                Op::f64_load => { body.push(0x2B); body.push(0x03); body.push(0x00); }
                Op::f64_store => { body.push(0x39); body.push(0x03); body.push(0x00); }

                Op::null => { body.push(0x42); write_leb128_i64(&mut body, 0); } // i64.const 0
                Op::r#true => { body.push(0x42); write_leb128_i64(&mut body, 1); }
                Op::r#false => { body.push(0x42); write_leb128_i64(&mut body, 0); }

                Op::r#const => {
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

                Op::call => { body.push(0x10); let argc = chunk.code[ip]; ip += 1; write_leb128_u32(&mut body, argc as u32); }
                Op::call_import => {
                    let import_idx = read_u16(&chunk.code, &mut ip);
                    let _argc = chunk.code[ip]; ip += 1;
                    body.push(0x10);
                    write_leb128_u32(&mut body, import_idx as u32);
                }

                Op::br => { let _ = read_i16(&chunk.code, &mut ip); body.push(0x0C); write_leb128_u32(&mut body, 0); }
                Op::br_if_false | Op::br_if_true => { let _ = read_i16(&chunk.code, &mut ip); body.push(0x0D); write_leb128_u32(&mut body, 0); }
                Op::block => { let _ = read_u16(&chunk.code, &mut ip); body.push(0x02); body.push(TYPE_VOID); }
                Op::r#loop => { let _ = read_u16(&chunk.code, &mut ip); body.push(0x03); body.push(TYPE_VOID); }
                Op::end => body.push(0x0B),

                // --- Dynamic ops → call vybe:rt import ---
                Op::dyn_add => emit_rt_call(&mut body, &rt_idx, "dyn_add"),
                Op::dyn_eq => emit_rt_call(&mut body, &rt_idx, "dyn_eq"),
                Op::dyn_ne => emit_rt_call(&mut body, &rt_idx, "dyn_ne"),
                Op::dyn_lt => emit_rt_call(&mut body, &rt_idx, "dyn_lt"),
                Op::dyn_gt => emit_rt_call(&mut body, &rt_idx, "dyn_gt"),
                Op::dyn_le => emit_rt_call(&mut body, &rt_idx, "dyn_le"),
                Op::dyn_ge => emit_rt_call(&mut body, &rt_idx, "dyn_ge"),
                Op::dyn_neg => emit_rt_call(&mut body, &rt_idx, "dyn_neg"),
                Op::dyn_not => emit_rt_call(&mut body, &rt_idx, "dyn_not"),
                Op::dyn_to_bool => emit_rt_call(&mut body, &rt_idx, "dyn_to_bool"),
                Op::str_concat => emit_rt_call(&mut body, &rt_idx, "str_concat"),

                Op::global_get => { let _ = read_u16(&chunk.code, &mut ip); emit_rt_call(&mut body, &rt_idx, "global_get"); }
                Op::global_set => { let _ = read_u16(&chunk.code, &mut ip); emit_rt_call(&mut body, &rt_idx, "global_set"); }

                Op::struct_get => { let _ = read_u16(&chunk.code, &mut ip); emit_rt_call(&mut body, &rt_idx, "get_prop"); }
                Op::struct_set => { let _ = read_u16(&chunk.code, &mut ip); emit_rt_call(&mut body, &rt_idx, "set_prop"); }
                Op::struct_new => { let _ = read_u16(&chunk.code, &mut ip); emit_rt_call(&mut body, &rt_idx, "new_object"); }
                Op::array_get => emit_rt_call(&mut body, &rt_idx, "array_get"),
                Op::array_set => emit_rt_call(&mut body, &rt_idx, "array_set"),
                Op::array_new => { let _ = read_u16(&chunk.code, &mut ip); emit_rt_call(&mut body, &rt_idx, "new_array"); }

                // --- Skip complex ops, emit nop ---
                Op::ref_func => { let _ = read_u16(&chunk.code, &mut ip); let uv = chunk.code[ip] as usize; ip += 1 + uv * 2; body.push(0x01); }
                Op::try_start => { ip += 4; body.push(0x01); }
                Op::upvalue_get | Op::upvalue_set => { ip += 1; body.push(0x01); }
                Op::str_concat_n | Op::return_call | Op::call_indirect | Op::pack | Op::br_label | Op::br_if_label => { ip += 1; body.push(0x01); }
                Op::br_table => { let count = chunk.code[ip] as usize; ip += 2 + count; body.push(0x01); }
                Op::ref_test => { ip += 2; body.push(0x01); }
                Op::dup => body.push(0x01), // nop (can't easily dup in WASM)

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
        let body_start = cpos;
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
        chunk.emit_op(Op::r#return, 0);
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

fn translate_wasm_to_chunk(wasm: &[u8], name: &str, arity: u8, wasm_local_count: u32, import_count: usize) -> Chunk {
    let mut chunk = Chunk::new(name);
    chunk.arity = arity;
    chunk.local_count = (1 + arity as u32 + wasm_local_count) as u16;

    let mut pos = 0;
    let mut label_stack: Vec<WasmLabel> = Vec::new();

    while pos < wasm.len() {
        let byte = wasm[pos]; pos += 1;

        match byte {
            0x00 => chunk.emit_op(Op::halt, 0),
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
                chunk.emit_op(Op::dyn_to_bool, 0);
                let patch = chunk.emit_jump(Op::br_if_false, 0);
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
                    let skip = chunk.emit_jump(Op::br, 0);
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
                    let end_offset = chunk.current_offset();
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
                        let patch = chunk.emit_jump(Op::br, 0);
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
                chunk.emit_op(Op::dyn_to_bool, 0);
                if let Some(label) = label_stack.iter().rev().nth(depth) {
                    if label.is_loop {
                        // Loop: conditional jump back
                        let exit = chunk.emit_jump(Op::br_if_false, 0);
                        chunk.emit_loop(label.start_offset, 0);
                        chunk.patch_jump(exit);
                    } else {
                        // Block: conditional jump forward
                        let patch = chunk.emit_jump(Op::br_if_true, 0);
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
                chunk.emit_op(Op::drop, 0); // simplified
            }
            0x0F => chunk.emit_op(Op::r#return, 0),
            0x1A => chunk.emit_op(Op::drop, 0),
            0x1B => {
                // select: [val1, val2, cond] → val1 if cond else val2
                // Implement as: if cond then drop val2 else drop val1
                // Simplified: just drop the condition and top value, keep bottom
                chunk.emit_op(Op::drop, 0); // drop cond — simplified select
            }

            // call — adjust index (skip imports, offset to our chunk indices)
            0x10 => {
                let (idx, _) = read_leb128_u32(&wasm[pos..]);
                skip_leb128(wasm, &mut pos);
                chunk.emit_op_u8(Op::call, idx as u8, 0);
            }

            // local.get — offset by 1 (our slot 0 = implicit fn)
            0x20 => {
                let (idx, _) = read_leb128_u32(&wasm[pos..]);
                skip_leb128(wasm, &mut pos);
                chunk.emit_op_u16(Op::local_get, (idx as u16) + 1, 0);
            }
            // local.set
            0x21 => {
                let (idx, _) = read_leb128_u32(&wasm[pos..]);
                skip_leb128(wasm, &mut pos);
                chunk.emit_op_u16(Op::local_set, (idx as u16) + 1, 0);
            }
            // local.tee
            0x22 => {
                let (idx, _) = read_leb128_u32(&wasm[pos..]);
                skip_leb128(wasm, &mut pos);
                chunk.emit_op_u16(Op::local_set, (idx as u16) + 1, 0);
            }

            // i32.const → add to constant pool as F64
            0x41 => {
                let (val, read) = read_leb128_i32(&wasm[pos..]);
                pos += read;
                let ci = chunk.add_constant(Value::F64(val as f64));
                chunk.emit_op_u16(Op::r#const, ci, 0);
            }

            // i64.const
            0x42 => {
                let (val, read) = read_leb128_i64(&wasm[pos..]);
                pos += read;
                let ci = chunk.add_constant(Value::F64(val as f64));
                chunk.emit_op_u16(Op::r#const, ci, 0);
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
                    chunk.emit_op_u16(Op::r#const, ci, 0);
                }
            }

            // f32.const
            0x43 => {
                if pos + 4 <= wasm.len() {
                    let val = f32::from_le_bytes([wasm[pos], wasm[pos+1], wasm[pos+2], wasm[pos+3]]);
                    pos += 4;
                    let ci = chunk.add_constant(Value::F64(val as f64));
                    chunk.emit_op_u16(Op::r#const, ci, 0);
                }
            }

            // i32 arithmetic — ALL opcodes
            0x67 => chunk.emit_op(Op::null, 0),      // i32.clz (approx)
            0x68 => chunk.emit_op(Op::null, 0),      // i32.ctz (approx)
            0x69 => chunk.emit_op(Op::null, 0),      // i32.popcnt (approx)
            0x6A => chunk.emit_op(Op::dyn_add, 0),   // i32.add
            0x6B => chunk.emit_op(Op::f64_sub, 0),   // i32.sub
            0x6C => chunk.emit_op(Op::f64_mul, 0),   // i32.mul
            0x6D => chunk.emit_op(Op::f64_div, 0),   // i32.div_s
            0x6E => chunk.emit_op(Op::f64_div, 0),   // i32.div_u
            0x6F => chunk.emit_op(Op::f64_mod, 0),   // i32.rem_s
            0x70 => chunk.emit_op(Op::f64_mod, 0),   // i32.rem_u
            0x71 => chunk.emit_op(Op::i32_and, 0),
            0x72 => chunk.emit_op(Op::i32_or, 0),
            0x73 => chunk.emit_op(Op::i32_xor, 0),
            0x74 => chunk.emit_op(Op::i32_shl, 0),
            0x75 => chunk.emit_op(Op::i32_shr_s, 0),
            0x76 => chunk.emit_op(Op::i32_shr_u, 0),
            0x77 => chunk.emit_op(Op::i32_shl, 0),   // i32.rotl (approx)
            0x78 => chunk.emit_op(Op::i32_shr_u, 0), // i32.rotr (approx)

            // i64 arithmetic
            0x7C => chunk.emit_op(Op::dyn_add, 0),   // i64.add
            0x7D => chunk.emit_op(Op::f64_sub, 0),   // i64.sub
            0x7E => chunk.emit_op(Op::f64_mul, 0),   // i64.mul
            0x7F => chunk.emit_op(Op::f64_div, 0),   // i64.div_s
            0x80 => chunk.emit_op(Op::f64_div, 0),   // i64.div_u
            0x81 => chunk.emit_op(Op::f64_mod, 0),   // i64.rem_s
            0x83 => chunk.emit_op(Op::i32_and, 0),   // i64.and
            0x84 => chunk.emit_op(Op::i32_or, 0),    // i64.or
            0x85 => chunk.emit_op(Op::i32_xor, 0),   // i64.xor
            0x86 => chunk.emit_op(Op::i32_shl, 0),   // i64.shl
            0x87 => chunk.emit_op(Op::i32_shr_s, 0), // i64.shr_s
            0x88 => chunk.emit_op(Op::i32_shr_u, 0), // i64.shr_u

            // i64 comparison
            0x50 => chunk.emit_op(Op::dyn_not, 0),   // i64.eqz
            0x51 => chunk.emit_op(Op::dyn_eq, 0),    // i64.eq
            0x52 => chunk.emit_op(Op::dyn_ne, 0),    // i64.ne
            0x53 => chunk.emit_op(Op::dyn_lt, 0),    // i64.lt_s
            0x55 => chunk.emit_op(Op::dyn_gt, 0),    // i64.gt_s
            0x57 => chunk.emit_op(Op::dyn_le, 0),    // i64.le_s
            0x59 => chunk.emit_op(Op::dyn_ge, 0),    // i64.ge_s

            // i32 comparison
            0x45 => chunk.emit_op(Op::dyn_not, 0),    // i32.eqz
            0x46 => chunk.emit_op(Op::dyn_eq, 0),     // i32.eq
            0x47 => chunk.emit_op(Op::dyn_ne, 0),     // i32.ne
            0x48 => chunk.emit_op(Op::dyn_lt, 0),     // i32.lt_s
            0x49 => chunk.emit_op(Op::dyn_lt, 0),     // i32.lt_u
            0x4A => chunk.emit_op(Op::dyn_gt, 0),     // i32.gt_s
            0x4B => chunk.emit_op(Op::dyn_gt, 0),     // i32.gt_u
            0x4C => chunk.emit_op(Op::dyn_le, 0),     // i32.le_s
            0x4D => chunk.emit_op(Op::dyn_le, 0),     // i32.le_u
            0x4E => chunk.emit_op(Op::dyn_ge, 0),     // i32.ge_s
            0x4F => chunk.emit_op(Op::dyn_ge, 0),     // i32.ge_u

            // f64 arithmetic
            0xA0 => chunk.emit_op(Op::f64_add, 0),
            0xA1 => chunk.emit_op(Op::f64_sub, 0),
            0xA2 => chunk.emit_op(Op::f64_mul, 0),
            0xA3 => chunk.emit_op(Op::f64_div, 0),

            // f64 comparison
            0x61 => chunk.emit_op(Op::dyn_eq, 0),
            0x62 => chunk.emit_op(Op::dyn_ne, 0),
            0x63 => chunk.emit_op(Op::dyn_lt, 0),
            0x64 => chunk.emit_op(Op::dyn_gt, 0),
            0x65 => chunk.emit_op(Op::dyn_le, 0),
            0x66 => chunk.emit_op(Op::dyn_ge, 0),

            // Memory
            0x28 => { skip_leb128(wasm, &mut pos); skip_leb128(wasm, &mut pos); chunk.emit_op(Op::i32_load, 0); }
            0x29 => { skip_leb128(wasm, &mut pos); skip_leb128(wasm, &mut pos); chunk.emit_op(Op::i64_load, 0); }
            0x2B => { skip_leb128(wasm, &mut pos); skip_leb128(wasm, &mut pos); chunk.emit_op(Op::f64_load, 0); }
            0x2D => { skip_leb128(wasm, &mut pos); skip_leb128(wasm, &mut pos); chunk.emit_op(Op::i32_load8_u, 0); }
            0x36 => { skip_leb128(wasm, &mut pos); skip_leb128(wasm, &mut pos); chunk.emit_op(Op::i32_store, 0); }
            0x37 => { skip_leb128(wasm, &mut pos); skip_leb128(wasm, &mut pos); chunk.emit_op(Op::i64_store, 0); }
            0x39 => { skip_leb128(wasm, &mut pos); skip_leb128(wasm, &mut pos); chunk.emit_op(Op::f64_store, 0); }
            0x3A => { skip_leb128(wasm, &mut pos); skip_leb128(wasm, &mut pos); chunk.emit_op(Op::i32_store8, 0); }
            0x3F => { skip_leb128(wasm, &mut pos); chunk.emit_op(Op::memory_size, 0); }
            0x40 => { skip_leb128(wasm, &mut pos); chunk.emit_op(Op::memory_grow, 0); }

            // f32 arithmetic (→ f64 in our VM)
            0x92 => chunk.emit_op(Op::f64_add, 0),   // f32.add
            0x93 => chunk.emit_op(Op::f64_sub, 0),   // f32.sub
            0x94 => chunk.emit_op(Op::f64_mul, 0),   // f32.mul
            0x95 => chunk.emit_op(Op::f64_div, 0),   // f32.div

            // f64 extra ops
            0x99 => chunk.emit_op(Op::dyn_neg, 0),   // f64.neg
            0x9A => chunk.emit_op(Op::null, 0),       // f64.ceil (approx)
            0x9B => chunk.emit_op(Op::null, 0),       // f64.floor (approx)
            0x9C => chunk.emit_op(Op::null, 0),       // f64.trunc (approx)
            0x9D => chunk.emit_op(Op::null, 0),       // f64.nearest (approx)
            0x9E => chunk.emit_op(Op::null, 0),       // f64.sqrt (approx)

            // ALL conversions
            0xA7 => {} // i32.wrap_i64
            0xA8 => chunk.emit_op(Op::i32_from_f64, 0), // i32.trunc_f32_s
            0xA9 => chunk.emit_op(Op::i32_from_f64, 0), // i32.trunc_f32_u
            0xAA => chunk.emit_op(Op::i32_from_f64, 0), // i32.trunc_f64_s
            0xAB => chunk.emit_op(Op::i32_from_f64, 0), // i32.trunc_f64_u
            0xAC => {} // i64.extend_i32_s
            0xAD => {} // i64.extend_i32_u
            0xAE => {} // i64.trunc_f32_s
            0xAF => {} // i64.trunc_f32_u
            0xB0 => {} // i64.trunc_f64_s
            0xB1 => {} // i64.trunc_f64_u
            0xB2 => chunk.emit_op(Op::f64_from_i32, 0), // f32.convert_i32_s
            0xB3 => chunk.emit_op(Op::f64_from_i32, 0), // f32.convert_i32_u
            0xB4 => {} // f32.convert_i64_s
            0xB5 => {} // f32.convert_i64_u
            0xB6 => {} // f32.demote_f64
            0xB7 => chunk.emit_op(Op::f64_from_i32, 0), // f64.convert_i32_s
            0xB8 => chunk.emit_op(Op::f64_from_i32, 0), // f64.convert_i32_u
            0xB9 => {} // f64.promote_f32
            0xBA => {} // f32.reinterpret_i32
            0xBB => {} // f64.reinterpret_i64
            0xBC => {} // i32.reinterpret_f32
            0xBD => {} // i64.reinterpret_f64

            // Sign extension
            0xC0 => {} // i32.extend8_s
            0xC1 => {} // i32.extend16_s
            0xC2 => {} // i64.extend8_s
            0xC3 => {} // i64.extend16_s
            0xC4 => {} // i64.extend32_s

            // global.get/set (WASM globals — we skip for now)
            0x23 => { skip_leb128(wasm, &mut pos); chunk.emit_op(Op::null, 0); }
            0x24 => { skip_leb128(wasm, &mut pos); chunk.emit_op(Op::drop, 0); }

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
        Value::Null => out.push(0),
        Value::Bool(b) => { out.push(1); out.push(if *b { 1 } else { 0 }); }
        Value::I32(n) => { out.push(2); out.extend_from_slice(&n.to_le_bytes()); }
        Value::I64(n) => { out.push(3); out.extend_from_slice(&n.to_le_bytes()); }
        Value::F64(n) => { out.push(4); out.extend_from_slice(&n.to_le_bytes()); }
        Value::String(s) => { out.push(5); write_name(out, s); }
        Value::Object(_) => out.push(0),
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
        5 => { let (len, read) = read_leb128_u32(&data[*pos..]); *pos += read; let s = std::str::from_utf8(&data[*pos..*pos+len as usize]).unwrap_or(""); *pos += len as usize; Value::String(Rc::from(s)) }
        _ => Value::Null,
    }
}

/// Calculate total bytes an opcode + operands consume in our internal format
fn opcode_size(op: Op, code: &[u8], ip: usize) -> usize {
    match op {
        Op::r#const | Op::local_get | Op::local_set | Op::global_get | Op::global_set
        | Op::struct_get | Op::struct_set | Op::struct_new | Op::array_new
        | Op::ref_test | Op::block | Op::r#loop | Op::class_new | Op::method_def
        | Op::canon_lift | Op::canon_lower => 3, // 1 op + 2 operand
        Op::br | Op::br_if_false | Op::br_if_true | Op::br_if_null => 3,
        Op::upvalue_get | Op::upvalue_set | Op::call | Op::str_concat_n
        | Op::return_call | Op::call_indirect | Op::pack | Op::br_label | Op::br_if_label => 2,
        Op::call_import => 4, // 1 + u16 + u8
        Op::try_start => 5, // 1 + u16 + u16
        Op::ref_func => {
            let uv_count = code.get(ip + 3).copied().unwrap_or(0) as usize;
            4 + uv_count * 2
        }
        Op::br_table => {
            let count = code.get(ip + 1).copied().unwrap_or(0) as usize;
            3 + count
        }
        _ => 1, // single byte opcodes
    }
}
