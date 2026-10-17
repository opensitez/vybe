//! Bounded unsigned-byte search in proven linear memory32 storage.
//! Callers own pointer/count validation and language-specific pointer results.
use vybe_runtime::chunk::{LocalName, LocalOrigin};
use vybe_runtime::{Chunk, Op};
fn get(c: &mut Chunk, s: u16) {
    c.emit_op_u16(Op::LOCAL_GET, s, 0);
}
fn set(c: &mut Chunk, s: u16) {
    c.emit_op_u16(Op::LOCAL_SET, s, 0);
}
/// Explicit linear memory32, numeric byte needle and a uint32 byte count.
/// Returns a Number byte index, or -1. Zero bytes are ordinary search data.
fn emit_body(c: &mut Chunk) {
    let b = c.alloc_scratch(6);
    let (ptr, needle, n, i, mask, splat) = (b, b + 1, b + 2, b + 3, b + 4, b + 5);
    // The effective address is dead before the SIMD bitmask is produced.
    let address = mask;
    c.local_names.push(LocalName::new(
        splat,
        "<native-wasm-scratch:v128>",
        LocalOrigin::Compiler,
    ));
    let cast = c.add_import("wasm:js-number", "toU32");
    c.emit_call(cast, 1, 0);
    set(c, n);
    c.emit_i32_const(255, 0);
    c.emit_op(Op::I32_AND, 0);
    set(c, needle);
    set(c, ptr);
    get(c, needle);
    c.emit_op(Op::I8X16_SPLAT, 0);
    set(c, splat);
    c.emit_i32_const(0, 0);
    set(c, i);
    let done = c.emit_block(0);
    let (repeat, _) = c.emit_loop_s(0);
    get(c, i);
    get(c, n);
    c.emit_op(Op::I32_GE_U, 0);
    c.emit_br_if(1, 0);
    // Try a scalar byte first; this also preserves immediate-hit latency.
    get(c, ptr);
    get(c, i);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_TEE, address, 0);
    c.emit_op(Op::I32_LOAD8_U, 0);
    get(c, needle);
    c.emit_op(Op::I32_EQ, 0);
    c.emit_if(0);
    get(c, i);
    c.emit_op(Op::F64_CONVERT_I32_U, 0);
    c.emit_op(Op::RETURN, 0);
    c.emit_end(0);
    get(c, n);
    get(c, i);
    c.emit_op(Op::I32_SUB, 0);
    c.emit_i32_const(16, 0);
    c.emit_op(Op::I32_GE_U, 0);
    c.emit_if(0);
    // A successful scalar load proves its complete Wasm page is mapped.
    get(c, address);
    c.emit_i32_const(65535, 0);
    c.emit_op(Op::I32_AND, 0);
    c.emit_i32_const(65520, 0);
    c.emit_op(Op::I32_LE_U, 0);
    c.emit_if(0);
    get(c, address);
    c.emit_op(Op::V128_LOAD, 0);
    c.emit_leb_u32(0x80, 0);
    c.emit_leb_u32(0, 0);
    get(c, splat);
    c.emit_op(Op::I8X16_EQ, 0);
    c.emit_op(Op::I8X16_BITMASK, 0);
    c.emit_op_u16(Op::LOCAL_TEE, mask, 0);
    c.emit_if(0);
    get(c, i);
    get(c, mask);
    c.emit_op(Op::I32_CTZ, 0);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op(Op::F64_CONVERT_I32_U, 0);
    c.emit_op(Op::RETURN, 0);
    c.emit_end(0);
    get(c, i);
    c.emit_i32_const(16, 0);
    c.emit_op(Op::I32_ADD, 0);
    set(c, i);
    c.emit_br(2, 0);
    c.emit_end(0);
    c.emit_end(0);
    get(c, i);
    c.emit_i32_const(1, 0);
    c.emit_op(Op::I32_ADD, 0);
    set(c, i);
    c.emit_br(0, 0);
    c.emit_end(0);
    c.patch_loop(repeat);
    c.emit_end(0);
    c.patch_block(done);
    c.emit_f64_const(-1.0, 0);
}

/// Consumes (address, numeric needle, uint32 count); produces Number index/-1.
/// Shares one standard-Wasm body per module.
pub fn emit_linear_index(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    const NAME: &str = "<primitive-memory32-byte-index>";
    let target = if let Some(index) = chunks.iter().position(|c| c.name == NAME) {
        index
    } else {
        let mut helper = Chunk::new(NAME);
        helper.arity = 3;
        helper.param_count = 3;
        helper.local_count = 3;
        helper.memory_min_pages = chunks[0].memory_min_pages.clone();
        helper.memory_is_64 = chunks[0].memory_is_64.clone();
        for slot in 0..3 {
            helper.emit_op_u16(Op::LOCAL_GET, slot, 0);
        }
        emit_body(&mut helper);
        helper.emit_op(Op::RETURN, 0);
        let index = chunks.len();
        chunks.push(helper);
        index
    };
    if super::functions::try_emit_direct_helper_call(chunks, current, target, 3, line) {
        return;
    }
    let helper = chunks[target].clone();
    chunks[current]
        .emit_direct_defined_call(
            u32::try_from(target).expect("Wasm function index"),
            &helper,
            false,
            line,
        )
        .expect("compiler-owned byte search has a direct Source signature");
}
