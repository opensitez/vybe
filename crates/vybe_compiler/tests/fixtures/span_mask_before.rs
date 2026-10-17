//! Single-byte membership spans in NUL-terminated linear memory32 strings.
use vybe_runtime::chunk::{LocalName, LocalOrigin};
use vybe_runtime::{Chunk, Op};

fn get(c: &mut Chunk, slot: u16) {
    c.emit_op_u16(Op::LOCAL_GET, slot, 0);
}
fn set(c: &mut Chunk, slot: u16) {
    c.emit_op_u16(Op::LOCAL_SET, slot, 0);
}
fn result(c: &mut Chunk, index: u16) {
    get(c, index);
    c.emit_op(Op::F64_CONVERT_I32_U, 0);
    c.emit_op(Op::RETURN, 0);
}
fn body(c: &mut Chunk, take_matches: bool) {
    let b = c.alloc_scratch(6);
    let (ptr, needle, index, byte, vector, splat) = (b, b + 1, b + 2, b + 3, b + 4, b + 5);
    for slot in [vector, splat] {
        c.local_names.push(LocalName::new(
            slot,
            "<native-wasm-scratch:v128>",
            LocalOrigin::Compiler,
        ));
    }
    c.emit_i32_const(255, 0);
    c.emit_op(Op::I32_AND, 0);
    set(c, needle);
    set(c, ptr);
    get(c, needle);
    c.emit_op(Op::I8X16_SPLAT, 0);
    set(c, splat);
    c.emit_i32_const(0, 0);
    set(c, index);
    let (repeat, _) = c.emit_loop_s(0);
    get(c, ptr);
    get(c, index);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op(Op::I32_LOAD8_U, 0);
    set(c, byte);
    // The scalar load proves the page exists. A vector stays within that page.
    get(c, ptr);
    get(c, index);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_i32_const(65535, 0);
    c.emit_op(Op::I32_AND, 0);
    c.emit_i32_const(65520, 0);
    c.emit_op(Op::I32_LE_U, 0);
    c.emit_if(0);
    get(c, ptr);
    get(c, index);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op(Op::V128_LOAD, 0);
    c.emit_leb_u32(0x80, 0);
    c.emit_leb_u32(0, 0);
    set(c, vector);
    get(c, vector);
    get(c, splat);
    c.emit_op(Op::I8X16_EQ, 0);
    c.emit_op(Op::I8X16_BITMASK, 0);
    if take_matches {
        // Needle is nonzero, so a NUL is also a nonmember.
        c.emit_i32_const(65535, 0);
        c.emit_op(Op::I32_XOR, 0);
    } else {
        get(c, vector);
        c.emit_op(Op::V128_CONST, 0);
        for _ in 0..16 {
            c.emit(0, 0);
        }
        c.emit_op(Op::I8X16_EQ, 0);
        c.emit_op(Op::I8X16_BITMASK, 0);
        c.emit_op(Op::I32_OR, 0);
    }
    c.emit_op_u16(Op::LOCAL_TEE, byte, 0);
    c.emit_if(0);
    get(c, index);
    get(c, byte);
    c.emit_op(Op::I32_CTZ, 0);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op(Op::F64_CONVERT_I32_U, 0);
    c.emit_op(Op::RETURN, 0);
    c.emit_end(0);
    get(c, index);
    c.emit_i32_const(16, 0);
    c.emit_op(Op::I32_ADD, 0);
    set(c, index);
    c.emit_br(1, 0);
    c.emit_end(0);
    get(c, byte);
    c.emit_op(Op::I32_EQZ, 0);
    c.emit_if(0);
    result(c, index);
    c.emit_end(0);
    get(c, byte);
    get(c, needle);
    c.emit_op(if take_matches { Op::I32_NE } else { Op::I32_EQ }, 0);
    c.emit_if(0);
    result(c, index);
    c.emit_end(0);
    get(c, index);
    c.emit_i32_const(1, 0);
    c.emit_op(Op::I32_ADD, 0);
    set(c, index);
    c.emit_br(0, 0);
    c.emit_end(0);
    c.patch_loop(repeat);
}

/// Consumes (address, nonzero byte), returns a Number byte length.
/// Language adapters prove linear storage and the single-byte set.
fn emit_shared_call(chunks: &mut Vec<Chunk>, current: usize, take_matches: bool, line: u32) {
    let name = if take_matches {
        "<primitive-memory32-span-member>"
    } else {
        "<primitive-memory32-span-nonmember>"
    };
    let target = if let Some(i) = chunks.iter().position(|c| c.name == name) {
        i
    } else {
        let mut c = Chunk::new(name);
        c.arity = 2;
        c.param_count = 2;
        c.local_count = 2;
        c.memory_min_pages = chunks[0].memory_min_pages.clone();
        c.memory_is_64 = chunks[0].memory_is_64.clone();
        for slot in 0..2 {
            get(&mut c, slot);
        }
        body(&mut c, take_matches);
        c.emit_f64_const(0.0, 0);
        c.emit_op(Op::RETURN, 0);
        let i = chunks.len();
        chunks.push(c);
        i
    };
    let helper = chunks[target].clone();
    chunks[current]
        .emit_direct_defined_call(
            u32::try_from(target).expect("Wasm function index"),
            &helper,
            false,
            line,
        )
        .expect("compiler-owned span has a direct Source signature");
}

/// Probe short spans inline before entering the shared vector body.
pub fn emit_single_byte(chunks: &mut Vec<Chunk>, current: usize, take_matches: bool, line: u32) {
    let c = &mut chunks[current];
    let b = c.alloc_scratch(3);
    let (ptr, needle, byte) = (b, b + 1, b + 2);
    c.emit_op_u16(Op::LOCAL_SET, needle, line);
    c.emit_op_u16(Op::LOCAL_SET, ptr, line);
    let mut branches = Vec::new();
    for offset in 0..2 {
        c.emit_op_u16(Op::LOCAL_GET, ptr, line);
        if offset != 0 {
            c.emit_i32_const(offset, line);
            c.emit_op(Op::I32_ADD, line);
        }
        c.emit_op(Op::I32_LOAD8_U, line);
        c.emit_op_u16(Op::LOCAL_SET, byte, line);
        c.emit_op_u16(Op::LOCAL_GET, byte, line);
        c.emit_op(Op::I32_EQZ, line);
        branches.push(c.emit_if_value(line));
        c.emit_f64_const(offset as f64, line);
        c.emit_else(line);
        c.emit_op_u16(Op::LOCAL_GET, byte, line);
        c.emit_op_u16(Op::LOCAL_GET, needle, line);
        c.emit_op(if take_matches { Op::I32_NE } else { Op::I32_EQ }, line);
        branches.push(c.emit_if_value(line));
        c.emit_f64_const(offset as f64, line);
        c.emit_else(line);
    }
    c.emit_op_u16(Op::LOCAL_GET, ptr, line);
    c.emit_i32_const(2, line);
    c.emit_op(Op::I32_ADD, line);
    c.emit_op_u16(Op::LOCAL_GET, needle, line);
    emit_shared_call(chunks, current, take_matches, line);
    let c = &mut chunks[current];
    c.emit_f64_const(2.0, line);
    c.emit_op(Op::F64_ADD, line);
    for branch in branches.into_iter().rev() {
        c.emit_end(line);
        c.patch_block(branch);
    }
}
