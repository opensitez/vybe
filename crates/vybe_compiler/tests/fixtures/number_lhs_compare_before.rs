use vybe_compiler::primitives::class_slots;
use vybe_runtime::{Chunk, Op};
#[derive(Clone, Copy)]
enum CmpOp {
    Lt,
}
fn alloc_locals(c: &mut Chunk, n: u16) -> u16 {
    c.alloc_scratch(n)
}
fn save(c: &mut Chunk, s: u16, l: u32) {
    c.emit_op_u16(Op::LOCAL_SET, s, l);
}
fn load(c: &mut Chunk, s: u16, l: u32) {
    c.emit_op_u16(Op::LOCAL_GET, s, l);
}
fn call1(c: &mut Chunk, i: u16, l: u32) {
    c.emit_call(i, 1, l);
}
fn call2(c: &mut Chunk, i: u16, l: u32) {
    c.emit_call(i, 2, l);
}
fn i32_const(c: &mut Chunk, v: i32, l: u32) {
    c.emit_i32_const(v, l);
}
fn f64_const(c: &mut Chunk, v: f64, l: u32) {
    c.emit_f64_const(v, l);
}
fn f64_cmp_op(_: &CmpOp) -> Op {
    Op::F64_LT
}
fn i32_cmp_op(_: &CmpOp) -> Op {
    Op::I32_LT_S
}
fn bigint_cmp_name(_: &CmpOp) -> &'static str {
    "lt"
}
fn emit_bigint_cmp(c: &mut Chunk, name: &str, l: u32) {
    let i = c.add_import("ecma:bigint", name);
    c.emit_call(i, 2, l);
    let i = c.add_import("wasm:js-boolean", "cast");
    c.emit_call(i, 1, l);
}
pub fn emit_lt(imports: &mut Chunk, c: &mut Chunk, line: u32) {
    emit_dyn_cmp_into(imports, c, line, CmpOp::Lt);
}
fn emit_object_field_to_slot(
    chunk: &mut Chunk,
    src_slot: u16,
    dst_slot: u16,
    field: &str,
    line: u32,
) {
    let field_key = class_slots::resolve_interned(
        chunk,
        &class_slots::ClassSlot::internal(field),
        &class_slots::PlainNames,
    );
    load(chunk, src_slot, line);
    class_slots::emit_class_get(
        chunk,
        class_slots::ObjSource::Stack,
        &field_key,
        class_slots::Dest::Stack,
        line,
    );
    save(chunk, dst_slot, line);
}
fn emit_slot_is_null_or_undefined(chunk: &mut Chunk, slot: u16, line: u32) {
    load(chunk, slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    load(chunk, slot, line);
    {
        let idx = chunk.add_import("wasm:js-undefined", "test");
        chunk.emit_call(idx, 1, line);
    }
    chunk.emit_op(Op::I32_OR, line);
}
fn emit_both_object_field_present(
    chunk: &mut Chunk,
    a_slot: u16,
    b_slot: u16,
    a_field_slot: u16,
    b_field_slot: u16,
    field: &str,
    line: u32,
) {
    emit_object_field_to_slot(chunk, a_slot, a_field_slot, field, line);
    emit_object_field_to_slot(chunk, b_slot, b_field_slot, field, line);
    emit_slot_is_null_or_undefined(chunk, a_field_slot, line);
    chunk.emit_op(Op::I32_EQZ, line);
    emit_slot_is_null_or_undefined(chunk, b_field_slot, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_op(Op::I32_AND, line);
}
fn emit_js_to_number_f64(
    chunk: &mut Chunk,
    slot: u16,
    to_f64: u16,
    test_bool: u16,
    cast_bool: u16,
    line: u32,
) {
    load(chunk, slot, line);
    {
        let idx = chunk.add_import("wasm:js-undefined", "test");
        chunk.emit_call(idx, 1, line);
    }
    chunk.emit_if_value(line);
    f64_const(chunk, f64::NAN, line);
    chunk.emit_else(line);

    load(chunk, slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if_value(line);
    f64_const(chunk, 0.0, line);
    chunk.emit_else(line);

    load(chunk, slot, line);
    call1(chunk, test_bool, line);
    chunk.emit_if_value(line);
    load(chunk, slot, line);
    call1(chunk, cast_bool, line);
    chunk.emit_op(Op::F64_FROM_I32, line);
    chunk.emit_else(line);
    load(chunk, slot, line);
    call1(chunk, to_f64, line);
    chunk.emit_end(line);

    chunk.emit_end(line);
    chunk.emit_end(line);
}
fn emit_dyn_cmp_into(_imports: &mut Chunk, code: &mut Chunk, line: u32, op: CmpOp) {
    let slots = alloc_locals(code, 4);
    let b_slot = slots;
    let a_slot = slots + 1;
    let b_time_slot = slots + 2;
    let a_time_slot = slots + 3;

    let test_num = code.add_import("wasm:js-number", "test");
    let to_f64 = code.add_import("wasm:js-number", "toF64");
    let test_str = code.add_import("wasm:js-string", "test");
    let str_compare = code.add_import("wasm:js-string", "compare");
    let test_bool = code.add_import("wasm:js-boolean", "test");
    let cast_bool = code.add_import("wasm:js-boolean", "cast");
    let test_bigint = code.add_import("wasm:js-bigint", "test");

    save(code, b_slot, line);
    save(code, a_slot, line);

    load(code, a_slot, line);
    call1(code, test_num, line);
    load(code, b_slot, line);
    call1(code, test_num, line);
    code.emit_op(Op::I32_AND, line);
    code.emit_if_i32(line);
    load(code, a_slot, line);
    call1(code, to_f64, line);
    load(code, b_slot, line);
    call1(code, to_f64, line);
    code.emit_op(f64_cmp_op(&op), line);
    code.emit_else(line);

    load(code, a_slot, line);
    call1(code, test_str, line);
    load(code, b_slot, line);
    call1(code, test_str, line);
    code.emit_op(Op::I32_AND, line);
    code.emit_if_i32(line);
    load(code, a_slot, line);
    load(code, b_slot, line);
    call2(code, str_compare, line);
    i32_const(code, 0, line);
    code.emit_op(i32_cmp_op(&op), line);
    code.emit_else(line);

    load(code, a_slot, line);
    call1(code, test_bigint, line);
    load(code, b_slot, line);
    call1(code, test_bigint, line);
    code.emit_op(Op::I32_AND, line);
    code.emit_if_i32(line);
    load(code, a_slot, line);
    load(code, b_slot, line);
    emit_bigint_cmp(code, bigint_cmp_name(&op), line);
    code.emit_else(line);
    emit_both_object_field_present(
        code,
        a_slot,
        b_slot,
        a_time_slot,
        b_time_slot,
        "Ticks",
        line,
    );
    code.emit_if_i32(line);
    load(code, a_time_slot, line);
    call1(code, to_f64, line);
    load(code, b_time_slot, line);
    call1(code, to_f64, line);
    code.emit_op(f64_cmp_op(&op), line);
    code.emit_else(line);
    emit_js_to_number_f64(code, a_slot, to_f64, test_bool, cast_bool, line);
    emit_js_to_number_f64(code, b_slot, to_f64, test_bool, cast_bool, line);
    code.emit_op(f64_cmp_op(&op), line);
    code.emit_end(line);
    code.emit_end(line);
    code.emit_end(line);
    code.emit_end(line);
    // Result is i32 — WASM-compliant
}
