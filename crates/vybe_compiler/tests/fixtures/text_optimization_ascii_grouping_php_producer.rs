// Captured PHP producer, only final grouping emission parameterized.
use std::sync::Arc;
use vybe_runtime::{Chunk, Op, Value};
fn alloc_local(chunk: &mut Chunk) -> u16 {
    chunk.alloc_scratch(1)
}
fn push_const(chunk: &mut Chunk, val: Value, line: u32) {
    match &val {
        Value::F64(v) => chunk.emit_f64_const(*v, line),
        Value::I32(v) => chunk.emit_i32_const(*v, line),
        Value::Null => chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line),
        Value::BigInt(v) => chunk.emit_i64_const(v.to_i64_wrapping(), line),
        Value::String(s) => chunk.emit_string_const(&s, line),
        Value::Bool(b) => chunk.emit_bool_const(*b, line),

        _ => {
            unreachable!("push_const: unexpected value type");
        }
    }
}
fn push_str(chunk: &mut Chunk, v: &str, line: u32) {
    push_const(chunk, Value::String(Arc::from(v)), line);
}
fn lset(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}
fn lget(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}
pub fn emit_number_format(
    chunks: &mut [Chunk],
    current: usize,
    argc: u8,
    line: u32,
    group: fn(&mut [Chunk], usize, u32),
) {
    let chunk = &mut chunks[current];
    let thousep_slot = alloc_local(chunk);
    let decsep_slot = alloc_local(chunk);
    let decimals_slot = alloc_local(chunk);
    let n_slot = alloc_local(chunk);
    let sign_slot = alloc_local(chunk);
    let fixed_slot = alloc_local(chunk);

    if argc >= 4 {
        lset(chunk, thousep_slot, line);
    } else {
        push_str(chunk, ",", line);
        lset(chunk, thousep_slot, line);
    }
    if argc >= 3 {
        lset(chunk, decsep_slot, line);
    } else {
        push_str(chunk, ".", line);
        lset(chunk, decsep_slot, line);
    }
    if argc >= 2 {
        lset(chunk, decimals_slot, line);
    } else {
        push_const(chunk, Value::F64(0.0), line);
        lset(chunk, decimals_slot, line);
    }

    // PHP 8: `number_format` decimals must be >= 0, else ValueError.
    lget(chunk, decimals_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    let _ = chunk;
    crate::emitter::type_guard::emit_throw_const(
        chunks,
        current,
        "ValueError",
        "number_format(): Argument #2 ($decimals) must be greater than or equal to 0",
        line,
    );
    let chunk = &mut chunks[current];
    chunk.emit_end(line);

    // n = +num (numeric coerce)
    push_const(chunk, Value::F64(0.0), line);
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
    lset(chunk, n_slot, line);

    // sign = ""; if n < 0: sign = "-"; n = -n
    push_str(chunk, "", line);
    lset(chunk, sign_slot, line);
    lget(chunk, n_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    push_str(chunk, "-", line);
    lset(chunk, sign_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    lget(chunk, n_slot, line);
    chunk.emit_op(Op::F64_SUB, line);
    lset(chunk, n_slot, line);
    chunk.emit_end(line);

    // PHP rounds half-away-from-zero; JS toFixed is engine-dependent
    // (banker's in V8, ECMA-spec deferred). Pre-round via Math.round
    // (which uses f64::round in vybe_host = half-away-from-zero) at
    // scale 10^decimals to force PHP semantics:
    //   n = Math.round(n * scale) / scale
    let pow = chunks[current].add_import("ecma:math", "pow");
    let round = chunks[current].add_import("ecma:math", "round");
    let chunk = &mut chunks[current];
    let scale_slot = alloc_local(chunk);
    push_const(chunk, Value::F64(10.0), line);
    lget(chunk, decimals_slot, line);
    chunk.emit_call(pow, 2, line);
    lset(chunk, scale_slot, line);
    lget(chunk, n_slot, line);
    lget(chunk, scale_slot, line);
    chunk.emit_op(Op::F64_MUL, line);
    chunk.emit_call(round, 1, line);
    lget(chunk, scale_slot, line);
    chunk.emit_op(Op::F64_DIV, line);
    lset(chunk, n_slot, line);

    // fixed = ecma:number.toFixed(n, decimals)
    lget(chunk, n_slot, line);
    lget(chunk, decimals_slot, line);
    let to_fixed = chunks[current].add_import("ecma:number", "toFixed");
    let chunk = &mut chunks[current];
    chunk.emit_call(to_fixed, 2, line);
    lset(chunk, fixed_slot, line);
    // Grouping is the SHARED emitter — python's `,` format flag binds the same
    // one. Everything above is php's and stays: the argument defaults, the
    // `ValueError` on negative decimals, and the half-away-from-zero pre-round
    // (JS `toFixed` is engine-dependent, so php's rounding cannot be delegated).
    //
    // The sign is re-attached AFTER grouping because `emit_group_digits`
    // preserves a leading sign itself; php stripped it before rounding, so it
    // is prepended here rather than passed through.
    lget(chunk, sign_slot, line);
    lget(chunk, fixed_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "concat");
        chunk.emit_call(idx, 2, line);
    }
    lget(chunk, thousep_slot, line);
    lget(chunk, decsep_slot, line);
    group(chunks, current, line);
}
