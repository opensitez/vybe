// Frozen JS relational emission before removing redundant Number extraction calls.
use vybe_runtime::{Chunk, Op};
fn call1(c: &mut Chunk, i: u16, l: u32) {
    c.emit_call(i, 1, l);
}
fn call2(c: &mut Chunk, i: u16, l: u32) {
    c.emit_call(i, 2, l);
}
fn save(c: &mut Chunk, i: u16, l: u32) {
    c.emit_op_u16(Op::LOCAL_SET, i, l);
}
fn load(c: &mut Chunk, i: u16, l: u32) {
    c.emit_op_u16(Op::LOCAL_GET, i, l);
}
fn alloc_locals(c: &mut Chunk, n: u16) -> u16 {
    c.alloc_scratch(n)
}
enum CmpOp {
    Lt,
    Gt,
    Le,
    Ge,
}

fn f64_cmp_op(op: &CmpOp) -> Op {
    match op {
        CmpOp::Lt => Op::F64_LT,
        CmpOp::Gt => Op::F64_GT,
        CmpOp::Le => Op::F64_LE,
        CmpOp::Ge => Op::F64_GE,
    }
}

fn i32_cmp_op(op: &CmpOp) -> Op {
    match op {
        CmpOp::Lt => Op::I32_LT_S,
        CmpOp::Gt => Op::I32_GT_S,
        CmpOp::Le => Op::I32_LE_S,
        CmpOp::Ge => Op::I32_GE_S,
    }
}

fn emit_bigint_cmp(chunk: &mut Chunk, name: &str, line: u32) {
    let op = chunk.add_import("ecma:bigint", name);
    call2(chunk, op, line); // [a, b] -> Bool
    let cast = chunk.add_import("wasm:js-boolean", "cast");
    call1(chunk, cast, line); // Bool -> i32, the ladder's arm type
}

fn bigint_cmp_name(op: &CmpOp) -> &'static str {
    match op {
        CmpOp::Lt => "lt",
        CmpOp::Gt => "gt",
        CmpOp::Le => "le",
        CmpOp::Ge => "ge",
    }
}

fn i32_const(chunk: &mut Chunk, v: i32, line: u32) {
    chunk.emit_i32_const(v, line);
}

fn emit_js_relational_cmp(
    chunk: &mut Chunk,
    line: u32,
    op: CmpOp,
    a_known_num: bool,
    b_known_num: bool,
) {
    // ── Operands this emitter MATERIALISED ITSELF ──────────────────────
    //
    // `a_known_num` / `b_known_num` mean the caller pushed that operand as a
    // numeric constant (`f64.const 13`), so every question the ladder asks
    // about it has a compile-time answer: `wasm:js-number.test` is 1,
    // `wasm:js-number.toF64` is the identity, and — because a constant number
    // is neither — `both string?` and `both bigint?` are FALSE whenever
    // EITHER operand is one. Asking anyway cost 2 host calls per iteration on
    // the taken path and 4 more on the mixed path, for `i < 13`.
    //
    // `Literal::BigInt` is a separate AST variant and is never reported here,
    // so the bigint arm is only ever dropped when it is provably unreachable.
    if a_known_num && b_known_num {
        // Both constants: the number arm is the only reachable one, and the
        // operands are already on the stack in order.
        chunk.emit_op(f64_cmp_op(&op), line);
        return;
    }
    if a_known_num || b_known_num {
        let slots = alloc_locals(chunk, 2);
        let b_slot = slots;
        let a_slot = slots + 1;

        let test_num = chunk.add_import("wasm:js-number", "test");
        let to_f64 = chunk.add_import("wasm:js-number", "toF64");
        let to_number = chunk.add_import("ecma:value", "toNumber");

        save(chunk, b_slot, line);
        save(chunk, a_slot, line);

        // Only the UNKNOWN side is asked; the `i32.and` goes with the other.
        load(chunk, if a_known_num { b_slot } else { a_slot }, line);
        call1(chunk, test_num, line);
        chunk.emit_if_i32(line);
        load(chunk, a_slot, line);
        if !a_known_num {
            call1(chunk, to_f64, line);
        }
        load(chunk, b_slot, line);
        if !b_known_num {
            call1(chunk, to_f64, line);
        }
        chunk.emit_op(f64_cmp_op(&op), line);
        chunk.emit_else(line);
        // Mixed: ToNumber the unknown side (NaN-safe — `"foo" < 0` is false).
        load(chunk, a_slot, line);
        if !a_known_num {
            call1(chunk, to_number, line);
        }
        load(chunk, b_slot, line);
        if !b_known_num {
            call1(chunk, to_number, line);
        }
        chunk.emit_op(f64_cmp_op(&op), line);
        chunk.emit_end(line);
        return;
    }

    let slots = alloc_locals(chunk, 2);
    let b_slot = slots;
    let a_slot = slots + 1;

    let test_num = chunk.add_import("wasm:js-number", "test");
    let to_f64 = chunk.add_import("wasm:js-number", "toF64");
    let test_str = chunk.add_import("wasm:js-string", "test");
    let str_compare = chunk.add_import("wasm:js-string", "compare");
    let test_bigint = chunk.add_import("wasm:js-bigint", "test");
    let to_number = chunk.add_import("ecma:value", "toNumber");

    save(chunk, b_slot, line);
    save(chunk, a_slot, line);

    // both number?
    load(chunk, a_slot, line);
    call1(chunk, test_num, line);
    load(chunk, b_slot, line);
    call1(chunk, test_num, line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_if_i32(line);
    load(chunk, a_slot, line);
    call1(chunk, to_f64, line);
    load(chunk, b_slot, line);
    call1(chunk, to_f64, line);
    chunk.emit_op(f64_cmp_op(&op), line);
    chunk.emit_else(line);

    // both string?
    load(chunk, a_slot, line);
    call1(chunk, test_str, line);
    load(chunk, b_slot, line);
    call1(chunk, test_str, line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_if_i32(line);
    load(chunk, a_slot, line);
    load(chunk, b_slot, line);
    call2(chunk, str_compare, line); // i32 (-1/0/1)
    i32_const(chunk, 0, line);
    chunk.emit_op(i32_cmp_op(&op), line);
    chunk.emit_else(line);

    // both bigint?
    load(chunk, a_slot, line);
    call1(chunk, test_bigint, line);
    load(chunk, b_slot, line);
    call1(chunk, test_bigint, line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_if_i32(line);
    load(chunk, a_slot, line);
    load(chunk, b_slot, line);
    emit_bigint_cmp(chunk, bigint_cmp_name(&op), line);
    chunk.emit_else(line);
    // fallback: ToNumber both (NaN-safe), compare as f64.
    load(chunk, a_slot, line);
    call1(chunk, to_number, line);
    load(chunk, b_slot, line);
    call1(chunk, to_number, line);
    chunk.emit_op(f64_cmp_op(&op), line);
    chunk.emit_end(line); // bigint
    chunk.emit_end(line); // string
    chunk.emit_end(line); // number
    // Result is i32 (0 or 1) — WASM-compliant for IF/BR_IF conditions.
}

pub fn emit_js_lt(chunk: &mut Chunk, line: u32, a_known_num: bool, b_known_num: bool) {
    emit_js_relational_cmp(chunk, line, CmpOp::Lt, a_known_num, b_known_num);
}
pub fn emit_js_gt(chunk: &mut Chunk, line: u32, a_known_num: bool, b_known_num: bool) {
    emit_js_relational_cmp(chunk, line, CmpOp::Gt, a_known_num, b_known_num);
}
pub fn emit_js_le(chunk: &mut Chunk, line: u32, a_known_num: bool, b_known_num: bool) {
    emit_js_relational_cmp(chunk, line, CmpOp::Le, a_known_num, b_known_num);
}
pub fn emit_js_ge(chunk: &mut Chunk, line: u32, a_known_num: bool, b_known_num: bool) {
    emit_js_relational_cmp(chunk, line, CmpOp::Ge, a_known_num, b_known_num);
}
