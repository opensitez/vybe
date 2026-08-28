//! PHP relational comparison — Rust inline opcode emitter.
//!
//! PHP `<` / `>` / `<=` / `>=` (and the `<=>` spaceship) compare two
//! strings lexicographically (`wasm:js-string.compare`) but fall back to the
//! numeric/dynamic comparison otherwise — unlike JS, which coerces to
//! primitive. DateTime objects are unboxed to their `__time` field
//! first so chronological comparison works.
//!
//! Mirrors the inline-emit shape of the other `languages/php/emitter`
//! adapters: writes WASM opcodes straight into the chunk, composing only
//! core ops + `vybe_compiler::primitives::ops` dynamic helpers. The shared compiler
//! routes here via the `string_aware_relational` profile flag — no
//! `profile.name == "php"` branch.

use std::sync::Arc;
use vybe_runtime::opcode::Op;
use vybe_runtime::{Chunk, Value};

use vybe_compiler::primitives::class_slots::{
    self, ClassSlot, Dest, ObjSource, PlainNames,
};
use vybe_compiler::primitives::ops::{emit_dyn_eq, emit_dyn_to_bool};

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
#[allow(dead_code)]
fn push_str(chunk: &mut Chunk, v: &str, line: u32) {
    push_const(chunk, Value::String(Arc::from(v)), line);
}
fn lset(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}
fn lget(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn emit_numeric_fallback(
    chunk: &mut Chunk,
    left_slot: u16,
    right_slot: u16,
    cmp_fn: fn(&mut Chunk, u32),
    line: u32,
) {
    let to_number = chunk.add_import("ecma:value", "toNumber");
    chunk.emit_op_u16(Op::LOCAL_GET, left_slot, line);
    chunk.emit_call(to_number, 1, line);
    chunk.emit_op_u16(Op::LOCAL_GET, right_slot, line);
    chunk.emit_call(to_number, 1, line);
    cmp_fn(chunk, line);
}

pub fn emit_php_loose_eq(chunks: &mut [Chunk], current: usize, _argc: u8, negate: bool, line: u32) {
    let parse_float = chunks[0].add_import("ecma:number", "parseFloat");
    let str_eq = chunks[0].add_import("wasm:js-string", "equals");
    let test_num = chunks[0].add_import("wasm:js-number", "test");
    let to_f64 = chunks[0].add_import("wasm:js-number", "toF64");
    let chunk = &mut chunks[current];
    let b_slot = alloc_local(chunk);
    let a_slot = alloc_local(chunk);
    let a_num_slot = alloc_local(chunk);
    let b_num_slot = alloc_local(chunk);

    lset(chunk, b_slot, line);
    lset(chunk, a_slot, line);

    lget(chunk, a_slot, line);
    let test_str_a = chunk.add_import("wasm:js-string", "test");
    chunk.emit_call(test_str_a, 1, line);
    chunk.emit_if_value(line);

    lget(chunk, b_slot, line);
    let test_str_b = chunk.add_import("wasm:js-string", "test");
    chunk.emit_call(test_str_b, 1, line);
    chunk.emit_if_value(line);

    lget(chunk, a_slot, line);
    chunk.emit_call(parse_float, 1, line);
    lset(chunk, a_num_slot, line);
    lget(chunk, b_slot, line);
    chunk.emit_call(parse_float, 1, line);
    lset(chunk, b_num_slot, line);

    lget(chunk, a_num_slot, line);
    lget(chunk, a_num_slot, line);
    chunk.emit_op(Op::F64_EQ, line);
    lget(chunk, b_num_slot, line);
    lget(chunk, b_num_slot, line);
    chunk.emit_op(Op::F64_EQ, line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_if_value(line);
    lget(chunk, a_num_slot, line);
    lget(chunk, b_num_slot, line);
    chunk.emit_op(Op::F64_EQ, line);
    vybe_compiler::primitives::ops::emit_i32_to_bool(chunk, line);
    chunk.emit_else(line);
    lget(chunk, a_slot, line);
    lget(chunk, b_slot, line);
    chunk.emit_call(str_eq, 2, line);
    vybe_compiler::primitives::ops::emit_i32_to_bool(chunk, line);
    chunk.emit_end(line);

    chunk.emit_else(line);
    lget(chunk, b_slot, line);
    chunk.emit_call(test_num, 1, line);
    chunk.emit_if_value(line);
    lget(chunk, a_slot, line);
    chunk.emit_call(parse_float, 1, line);
    lget(chunk, b_slot, line);
    chunk.emit_call(to_f64, 1, line);
    chunk.emit_op(Op::F64_EQ, line);
    vybe_compiler::primitives::ops::emit_i32_to_bool(chunk, line);
    chunk.emit_else(line);
    lget(chunk, a_slot, line);
    lget(chunk, b_slot, line);
    emit_dyn_eq(chunk, line);
    vybe_compiler::primitives::ops::emit_i32_to_bool(chunk, line);
    chunk.emit_end(line);
    chunk.emit_end(line);

    chunk.emit_else(line);
    lget(chunk, b_slot, line);
    chunk.emit_call(test_str_b, 1, line);
    chunk.emit_if_value(line);
    lget(chunk, a_slot, line);
    chunk.emit_call(test_num, 1, line);
    chunk.emit_if_value(line);
    lget(chunk, a_slot, line);
    chunk.emit_call(to_f64, 1, line);
    lget(chunk, b_slot, line);
    chunk.emit_call(parse_float, 1, line);
    chunk.emit_op(Op::F64_EQ, line);
    vybe_compiler::primitives::ops::emit_i32_to_bool(chunk, line);
    chunk.emit_else(line);
    lget(chunk, a_slot, line);
    lget(chunk, b_slot, line);
    emit_dyn_eq(chunk, line);
    vybe_compiler::primitives::ops::emit_i32_to_bool(chunk, line);
    chunk.emit_end(line);
    chunk.emit_else(line);
    lget(chunk, a_slot, line);
    lget(chunk, b_slot, line);
    emit_dyn_eq(chunk, line);
    vybe_compiler::primitives::ops::emit_i32_to_bool(chunk, line);
    chunk.emit_end(line);
    chunk.emit_end(line);

    if negate {
        emit_dyn_to_bool(chunk, line);
        chunk.emit_op(Op::I32_EQZ, line);
        vybe_compiler::primitives::ops::emit_i32_to_bool(chunk, line);
    }
}

/// Consume the top two stack values (`[a, b]`) and push `a <op> b` using
/// PHP comparison semantics, where `cmp_fn` emits the numeric/dynamic
/// fallback op (e.g. `emit_dyn_lt`).
pub fn emit_relational_compare(chunk: &mut Chunk, cmp_fn: fn(&mut Chunk, u32), line: u32) {
    let t_b = alloc_local(chunk);
    let t_a = alloc_local(chunk);
    let a_num = alloc_local(chunk);
    let b_num = alloc_local(chunk);
    chunk.emit_op_u16(Op::LOCAL_SET, t_b, line);
    chunk.emit_op_u16(Op::LOCAL_SET, t_a, line);

    maybe_unbox_datetime(chunk, t_a, line);
    maybe_unbox_datetime(chunk, t_b, line);

    chunk.emit_op_u16(Op::LOCAL_GET, t_a, line);
    let test_str_ta = chunk.add_import("wasm:js-string", "test");
    chunk.emit_call(test_str_ta, 1, line);
    chunk.emit_if_value(line);

    chunk.emit_op_u16(Op::LOCAL_GET, t_b, line);
    let test_str_tb = chunk.add_import("wasm:js-string", "test");
    chunk.emit_call(test_str_tb, 1, line);
    chunk.emit_if_value(line);

    // Both strings: PHP compares numerically when both are numeric strings,
    // otherwise lexicographically.
    let parse_float = chunk.add_import("ecma:number", "parseFloat");
    chunk.emit_op_u16(Op::LOCAL_GET, t_a, line);
    chunk.emit_call(parse_float, 1, line);
    chunk.emit_op_u16(Op::LOCAL_SET, a_num, line);
    chunk.emit_op_u16(Op::LOCAL_GET, t_b, line);
    chunk.emit_call(parse_float, 1, line);
    chunk.emit_op_u16(Op::LOCAL_SET, b_num, line);

    // NaN is the only f64 not equal to itself.
    chunk.emit_op_u16(Op::LOCAL_GET, a_num, line);
    chunk.emit_op_u16(Op::LOCAL_GET, a_num, line);
    chunk.emit_op(Op::F64_EQ, line);
    chunk.emit_op_u16(Op::LOCAL_GET, b_num, line);
    chunk.emit_op_u16(Op::LOCAL_GET, b_num, line);
    chunk.emit_op(Op::F64_EQ, line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::LOCAL_GET, a_num, line);
    chunk.emit_op_u16(Op::LOCAL_GET, b_num, line);
    cmp_fn(chunk, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, t_a, line);
    chunk.emit_op_u16(Op::LOCAL_GET, t_b, line);
    {
        let idx = chunk.add_import("wasm:js-string", "compare");
        chunk.emit_call(idx, 2, line);
    }
    push_const(chunk, Value::I32(0), line);
    cmp_fn(chunk, line);
    chunk.emit_end(line);

    chunk.emit_else(line);
    emit_numeric_fallback(chunk, t_a, t_b, cmp_fn, line);
    chunk.emit_end(line);

    chunk.emit_else(line);
    emit_numeric_fallback(chunk, t_a, t_b, cmp_fn, line);
    chunk.emit_end(line);
}

/// PHP's `<=>` — the THREE-WAY compare, `[a, b]` → `i32` in `{-1, 0, 1}`.
///
/// This is the primitive. `<`, `<=`, `>`, `>=` are its sign against 0
/// (builtinslotplan.md §2f), which is the inverse of how the shared emitter
/// used to work: `Spaceship` was built by calling the relational hook TWICE,
/// once for `<` and once for `>`. Spaceship is its own operator, not three.
///
/// Reached as the emit target `common:php.compare3`, declared by PHP's profile
/// as `[builtin_slots.string] compare`. No `LanguageHooks` callback and no
/// `profile.name` check is involved.
///
/// # PHP 8 comparison semantics, and the bug this fixes
///
/// PHP 8 compares two operands numerically **only when both are numeric**;
/// otherwise the number is cast to a string and the comparison is
/// lexicographic. `emit_relational_compare` tested numericness with
/// `ecma:number.parseFloat`, and `parseFloat("9a")` is `9`, not `NaN` — so
/// `"9a"` was treated as the number 9. Measured against the real `php` binary
/// 2026-07-31:
///
/// | expression | php | old vybe |
/// |---|---|---|
/// | `"10" <=> "9a"` | -1 | **1** |
/// | `5 <=> "abc"` | -1 | **0** |
/// | `0 <=> "abc"` | -1 | **0** |
/// | `"10" < "9a"` | true | **false** |
/// | `5 < "abc"` | true | **false** |
///
/// `ecma:value.toNumber` is the right test: it bottoms out on
/// `Value::as_f64`, which does a whole-string `parse::<f64>()` and yields NaN
/// for `"9a"` — exactly PHP's `is_numeric`, including the leading-whitespace
/// tolerance that makes `" 1" == "1"` true.
pub fn emit_compare3(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let t_b = alloc_local(chunk);
    let t_a = alloc_local(chunk);
    let a_num = alloc_local(chunk);
    let b_num = alloc_local(chunk);
    let res = alloc_local(chunk);
    lset(chunk, t_b, line);
    lset(chunk, t_a, line);

    // DateTime operands compare chronologically, via their `__time` field.
    maybe_unbox_datetime(chunk, t_a, line);
    maybe_unbox_datetime(chunk, t_b, line);

    // ── PHP 8 comparison table, row 1: `null` against a `string` ──────────
    // NULL becomes `""` and the ordinary string rules take over. This is why
    // `null <=> "0"` is -1 (lexical, `"" < "0"`) and not 0, which is what the
    // bool rule below would give — `(bool)null` and `(bool)"0"` are both false.
    coerce_null_against_string(chunk, t_a, t_b, line);
    coerce_null_against_string(chunk, t_b, t_a, line);

    // ── row 2: `bool` or `null` on EITHER side → compare both as bools ────
    // Reached only after row 1, so a null paired with a string is already a
    // string by now. Measured against the real `php` binary: `true <=> "abc"`
    // is 0 under this rule, where a string compare would say `"1" <=> "abc"`.
    lget(chunk, t_a, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    lget(chunk, t_a, line);
    let test_bool_a = chunk.add_import("wasm:js-boolean", "test");
    chunk.emit_call(test_bool_a, 1, line);
    chunk.emit_op(Op::I32_OR, line);
    lget(chunk, t_b, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_op(Op::I32_OR, line);
    lget(chunk, t_b, line);
    let test_bool_b = chunk.add_import("wasm:js-boolean", "test");
    chunk.emit_call(test_bool_b, 1, line);
    chunk.emit_op(Op::I32_OR, line);
    chunk.emit_if_value(line);

    // PHP falsiness, via PHP's own `empty` — NOT `ops::emit_dyn_to_bool`,
    // which is JS's and calls `"0"` truthy. Each side lands as an i32 in
    // {0, 1}, so the difference IS the three-way result: FALSE < TRUE.
    emit_php_truthy(chunks, current, t_a, line);
    emit_php_truthy(chunks, current, t_b, line);
    let chunk = &mut chunks[current];
    chunk.emit_op(Op::I32_SUB, line);

    chunk.emit_else(line);

    let to_number = chunk.add_import("ecma:value", "toNumber");
    lget(chunk, t_a, line);
    chunk.emit_call(to_number, 1, line);
    lset(chunk, a_num, line);
    lget(chunk, t_b, line);
    chunk.emit_call(to_number, 1, line);
    lset(chunk, b_num, line);

    // Both numeric? NaN is the only f64 not equal to itself.
    lget(chunk, a_num, line);
    lget(chunk, a_num, line);
    chunk.emit_op(Op::F64_EQ, line);
    lget(chunk, b_num, line);
    lget(chunk, b_num, line);
    chunk.emit_op(Op::F64_EQ, line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_if_value(line);

    // Numeric: (a > b) - (a < b).
    lget(chunk, a_num, line);
    lget(chunk, b_num, line);
    chunk.emit_op(Op::F64_GT, line);
    lget(chunk, a_num, line);
    lget(chunk, b_num, line);
    chunk.emit_op(Op::F64_LT, line);
    chunk.emit_op(Op::I32_SUB, line);

    chunk.emit_else(line);

    // Lexicographic. `ecma:string.String` is the string cast — the same host
    // function `primitives::strings::emit_to_string` uses, so a number operand
    // becomes its PHP string form. `wasm:js-string.concat` will NOT do this:
    // it rejects a non-string argument outright.
    let to_str = chunk.add_import("ecma:string", "String");
    lget(chunk, t_a, line);
    chunk.emit_call(to_str, 1, line);
    lget(chunk, t_b, line);
    chunk.emit_call(to_str, 1, line);
    let compare = chunk.add_import("wasm:js-string", "compare");
    chunk.emit_call(compare, 2, line);
    // `compare` is only documented by SIGN, so normalise to -1/0/1 rather than
    // letting a host implementation's magnitude leak into `<=>`'s result.
    lset(chunk, res, line);
    lget(chunk, res, line);
    chunk.emit_i32_const(0, line);
    chunk.emit_op(Op::I32_GT_S, line);
    lget(chunk, res, line);
    chunk.emit_i32_const(0, line);
    chunk.emit_op(Op::I32_LT_S, line);
    chunk.emit_op(Op::I32_SUB, line);

    chunk.emit_end(line); // numeric / lexical
    chunk.emit_end(line); // bool-or-null / everything else
}

/// PHP truthiness of the value in `slot`, pushed as an i32 in {0, 1}.
///
/// `!empty($v)`, which is exactly how the walker spells a PHP condition
/// (`php_truthy_condition`). Deliberately not `ops::emit_dyn_to_bool`: that is
/// JS truthiness, under which `"0"` and `[]` are both true and PHP says false.
fn emit_php_truthy(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    super::array_adapter::emit_php_empty_from_slot(chunks, current, slot, line);
    let chunk = &mut chunks[current];
    emit_dyn_to_bool(chunk, line);
    chunk.emit_op(Op::I32_EQZ, line);
}

/// PHP 8 comparison table row 1 — `null` compared against a `string` compares
/// `""` against that string. Rewrites `slot` in place when it holds null and
/// `other` holds a string, so the general legs below never see the null.
fn coerce_null_against_string(chunk: &mut Chunk, slot: u16, other: u16, line: u32) {
    lget(chunk, slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    lget(chunk, other, line);
    let test_str = chunk.add_import("wasm:js-string", "test");
    chunk.emit_call(test_str, 1, line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_if(line);
    push_str(chunk, "", line);
    lset(chunk, slot, line);
    chunk.emit_end(line);
}

/// If the value in `slot` is a boxed DateTime-like object, replace it
/// with its `__time` field so comparisons operate on the timestamp.
fn maybe_unbox_datetime(chunk: &mut Chunk, slot: u16, line: u32) {
    // object test: not null AND not number AND not string AND not boolean
    let obj_dt_slot = alloc_local(chunk);
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, obj_dt_slot, line);
    // not null
    chunk.emit_op_u16(Op::LOCAL_GET, obj_dt_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_op(Op::I32_EQZ, line);
    // AND not number
    chunk.emit_op_u16(Op::LOCAL_GET, obj_dt_slot, line);
    let test_num_dt = chunk.add_import("wasm:js-number", "test");
    chunk.emit_call(test_num_dt, 1, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_op(Op::I32_AND, line);
    // AND not string
    chunk.emit_op_u16(Op::LOCAL_GET, obj_dt_slot, line);
    let test_str_dt = chunk.add_import("wasm:js-string", "test");
    chunk.emit_call(test_str_dt, 1, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_op(Op::I32_AND, line);
    // AND not boolean
    chunk.emit_op_u16(Op::LOCAL_GET, obj_dt_slot, line);
    let test_bool_dt = chunk.add_import("wasm:js-boolean", "test");
    chunk.emit_call(test_bool_dt, 1, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_if(line);

    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    let cs_slot = class_slots::resolve(&ClassSlot::Internal(("__time").to_string()), &PlainNames);
    class_slots::emit_class_get(chunk, ObjSource::Stack, &cs_slot, Dest::Stack, line);
    let time_slot = alloc_local(chunk);
    chunk.emit_op_u16(Op::LOCAL_SET, time_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, time_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::LOCAL_GET, time_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
    chunk.emit_end(line);
    chunk.emit_end(line);
}
