// Separate typed control: only scalar nonempty IF result is repaired.
// Frozen actual PHP emitter and exact prerequisite coercion implementation.
//! PHP string helpers — Rust inline opcode emitters.
//!
//! Each `emit_*` writes opcodes directly into `chunks[current]`,
//! composing only WASM string ops (`wasm:js-string.length`, `wasm:js-string.charCodeAt`,
//! `ecma:string.indexOf`, `STR_TO_LOWER`, etc.) and where ECMA-262 already
//! covers the surface, `ecma:string.{encodeURIComponent,
//! decodeURIComponent}` / `ecma:number.toFixed`. No PHP-specific
//! host fns; no JS polyfills.

use std::sync::Arc;
use vybe_compiler::primitives::class_slots::{
    self, ClassSlot, Dest, ObjSource, PlainNames, ValueSource,
};
use vybe_compiler::primitives::{string_encoding, string_similarity, strings, url};
use vybe_runtime::opcode::Op;
use vybe_runtime::{Chunk, Value};

// ── Local-slot / push helpers (mirror datetime_adapter) ────────────

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
/// PHP's value → string coercion, for the ~55 ADAPTER call sites
/// (`strlen($v)`, `str_pad($v, …)`, `explode(…, $v)`, …).
///
/// This is the SAME emitter `.` concatenation and `echo` use, and the same one
/// `[builtin_slots.string] to_string` binds. It used to be `"" + v` through
/// `emit_dyn_add` — a THIRD implementation, which disagreed with the other two.
/// Measured against real php on `[true, false, null, 1.0, 1.5, 100, "x"]`:
///
/// | | result |
/// |---|---|
/// | real php `strlen($v)` | `1,0,0,1,3,3,1` |
/// | `coerce_to_str` (was) | `1,1,1,1,3,3,1` |
/// | `strlen("" . $v)` — the BOUND path | `1,0,0,1,3,3,1` |
///
/// So `strlen(false)` and `strlen(null)` were wrong while the declared binding
/// was already right: the fork, not the binding, was the bug.
fn coerce_to_str(chunk: &mut Chunk, line: u32) {
    emit_stringify(chunk, line);
}

fn call_import(
    chunks: &mut [Chunk],
    current: usize,
    module: &str,
    name: &str,
    argc: u8,
    line: u32,
) {
    let idx = chunks[current].add_import(module.to_string(), name.to_string());
    let chunk = &mut chunks[current];
    chunk.emit_call(idx, argc, line);
}

fn emit_stringify(chunk: &mut Chunk, line: u32) {
    let test_bigint = chunk.add_import("wasm:js-bigint", "test");
    let bigint_to_string = chunk.add_import("ecma:bigint", "toString");
    let from_i64 = chunk.add_import("wasm:js-string", "fromI64");
    let test_str = chunk.add_import("wasm:js-string", "test");
    let v_slot = alloc_local(chunk);
    let _ty_slot = alloc_local(chunk);

    lset(chunk, v_slot, line);

    // ── Already a string: hand it straight back ────────────────────────────
    //
    // NOT a shortcut with a different answer — the SAME answer with the ladder
    // removed. A string falls through every branch below (`REF_IS_NULL` false;
    // not boolean, bigint or number; `v !== v` false because a string equals
    // itself, which is why the INF/NAN probes cannot catch the literal strings
    // `"INF"`/`"NAN"`) and lands on the terminal default, `"" + v` — and `"" + s`
    // is the identity on a string. The comment on the NaN probe below already
    // says as much.
    //
    // Worth the test because php reaches here from THREE positions — `.`,
    // `echo`/`print`, and interpolation `"$a$b"` — and in all three the operand
    // is usually already a string. The ladder it skips runs three `emit_dyn_eq`
    // ladders and an `emit_dyn_add` ladder: 556 executed instructions per `.`,
    // measured.
    //
    // Objects stay on the PHP __toString dispatch path below.
    lget(chunk, v_slot, line);
    chunk.emit_call(test_str, 1, line);
    chunk.emit_if_value(line);
    lget(chunk, v_slot, line);
    chunk.emit_else(line);

    lget(chunk, v_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if_value(line);
    push_str(chunk, "", line);
    chunk.emit_else(line);

    // boolean test
    lget(chunk, v_slot, line);
    let test_bool_echo = chunk.add_import("wasm:js-boolean", "test");
    chunk.emit_call(test_bool_echo, 1, line);
    chunk.emit_if_value(line);
    lget(chunk, v_slot, line);
    let cast_bool_echo = chunk.add_import("wasm:js-boolean", "cast");
    chunk.emit_call(cast_bool_echo, 1, line);
    chunk.emit_if_value(line);
    push_str(chunk, "1", line);
    chunk.emit_else(line);
    push_str(chunk, "", line);
    chunk.emit_end(line);
    chunk.emit_else(line);
    // Not boolean. PHP integer boundary constants use BigInt internally so
    // they retain all 64-bit digits; stringify them through ECMA BigInt.
    lget(chunk, v_slot, line);
    chunk.emit_call(test_bigint, 1, line);
    chunk.emit_if_value(line);
    lget(chunk, v_slot, line);
    chunk.emit_call(bigint_to_string, 1, line);
    chunk.emit_else(line);
    // Not bigint. Keep non-numbers on the old `"" + v` path below so object
    // coercion and `__toString` stay centralized there, but use direct number
    // slots for actual numbers. The previous shape used three full dynamic
    // equality ladders (`v !== v`, `v == INF`, `v == -INF`) for every unknown
    // concat operand; WordPress path constants hit this thousands of times.
    lget(chunk, v_slot, line);
    let num_test_echo = chunk.add_import("wasm:js-number", "test");
    chunk.emit_call(num_test_echo, 1, line);
    chunk.emit_if_value(line);

    let n_slot = alloc_local(chunk);
    let to_f64_echo = chunk.add_import("wasm:js-number", "toF64");
    lget(chunk, v_slot, line);
    chunk.emit_call(to_f64_echo, 1, line);
    lset(chunk, n_slot, line);

    // NaN?  (n !== n)
    lget(chunk, n_slot, line);
    lget(chunk, n_slot, line);
    chunk.emit_op(Op::F64_NE, line);
    chunk.emit_if_value(line);
    push_str(chunk, "NAN", line);
    chunk.emit_else(line);

    // +Infinity?
    lget(chunk, n_slot, line);
    push_const(chunk, Value::F64(f64::INFINITY), line);
    chunk.emit_op(Op::F64_EQ, line);
    chunk.emit_if_value(line);
    push_str(chunk, "INF", line);
    chunk.emit_else(line);

    // -Infinity?
    lget(chunk, n_slot, line);
    push_const(chunk, Value::F64(f64::NEG_INFINITY), line);
    chunk.emit_op(Op::F64_EQ, line);
    chunk.emit_if_value(line);
    push_str(chunk, "-INF", line);
    chunk.emit_else(line);

    // Finite number. PHP stringifies floats with `precision=14` significant digits
    // (php_gcvt), not the VM's shortest-round-trip form. Normalize a float via
    // toPrecision(14) then parseFloat — the round-trip trims trailing zeros
    // exactly like PHP (`1.4142135623730951`→`1.4142135623731`, `0.1+0.2`→
    // `0.3`). Integers use ECMA String(), which matches the old `"" + v`
    // result without emitting the dynamic-add ladder.
    let to_string_echo = chunk.add_import("ecma:string", "String");
    lget(chunk, v_slot, line);
    let is_int_echo = chunk.add_import("ecma:number", "isInteger");
    chunk.emit_call(is_int_echo, 1, line);
    let cast_bool_number = chunk.add_import("wasm:js-boolean", "cast");
    chunk.emit_call(cast_bool_number, 1, line);
    chunk.emit_if_value(line);
    // integer-valued number.
    lget(chunk, v_slot, line);
    chunk.emit_call(to_string_echo, 1, line);
    chunk.emit_else(line);
    // fractional float: String(parseFloat(toPrecision(v, 14)))
    lget(chunk, v_slot, line);
    push_const(chunk, Value::I32(14), line);
    let to_prec_echo = chunk.add_import("ecma:number", "toPrecision");
    chunk.emit_call(to_prec_echo, 2, line);
    let parse_f_echo = chunk.add_import("ecma:number", "parseFloat");
    chunk.emit_call(parse_f_echo, 1, line);
    chunk.emit_call(to_string_echo, 1, line);
    chunk.emit_end(line); // integer / fractional
    chunk.emit_end(line); // -Infinity / finite
    chunk.emit_end(line); // +Infinity
    chunk.emit_end(line); // NaN
    chunk.emit_else(line); // non-number
                           // Keep the DOM/SimpleXML text conversion used by explicit PHP casts.
    let text_key = vybe_compiler::primitives::class_slots::resolve(
        &vybe_compiler::primitives::class_slots::ClassSlot::internal("textContent"),
        &vybe_compiler::primitives::class_slots::PlainNames,
    );
    vybe_compiler::primitives::class_slots::emit_class_get(
        chunk,
        vybe_compiler::primitives::class_slots::ObjSource::Local(v_slot),
        &text_key,
        vybe_compiler::primitives::class_slots::Dest::Stack,
        line,
    );
    let text_slot = alloc_local(chunk);
    lset(chunk, text_slot, line);
    lget(chunk, text_slot, line);
    chunk.emit_call(test_str, 1, line);
    chunk.emit_if_value(line);
    lget(chunk, text_slot, line);
    chunk.emit_else(line);
    // PHP object conversion invokes __toString, including methods installed
    // by runtime includes. ECMAScript ToPrimitive does not know that name.
    let method_key = vybe_compiler::primitives::class_slots::resolve(
        &vybe_compiler::primitives::class_slots::ClassSlot::Slot(vybe_ast::ProtocolSlot::ToString),
        &vybe_compiler::primitives::class_slots::PlainNames,
    );
    vybe_compiler::primitives::class_slots::emit_class_get(
        chunk,
        vybe_compiler::primitives::class_slots::ObjSource::Local(v_slot),
        &method_key,
        vybe_compiler::primitives::class_slots::Dest::Stack,
        line,
    );
    let method_slot = alloc_local(chunk);
    lset(chunk, method_slot, line);
    lget(chunk, method_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if_value(line);
    push_str(chunk, "", line);
    lget(chunk, v_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
    chunk.emit_else(line);
    lget(chunk, method_slot, line);
    lget(chunk, v_slot, line);
    vybe_compiler::primitives::callable::emit_direct_invoke_chunk(chunk, 1, line);
    strings::emit_to_string(chunk, line);
    chunk.emit_end(line);
    chunk.emit_end(line); // DOM text / magic method
    chunk.emit_end(line); // number / object
    chunk.emit_end(line); // bigint
    chunk.emit_end(line); // boolean
    chunk.emit_end(line); // null
    chunk.emit_end(line); // string
}

fn emit_str_replace_count_occurrences(
    chunks: &mut [Chunk],
    current: usize,
    subj_slot: u16,
    needle_slot: u16,
    count_slot: u16,
    line: u32,
) {
    {
        let chunk = &mut chunks[current];
        lget(chunk, subj_slot, line);
        lget(chunk, needle_slot, line);
    }
    call_import(chunks, current, "ecma:string", "split", 2, line);
    let chunk = &mut chunks[current];
    chunk.emit_op(Op::ARRAY_LENGTH, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_SUB, line);
    lget(chunk, count_slot, line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, count_slot, line);
}

fn emit_class_get_internal(chunk: &mut Chunk, key: &str, line: u32) {
    let slot = class_slots::resolve(&ClassSlot::internal(key), &PlainNames);
    class_slots::emit_class_get(chunk, ObjSource::Stack, &slot, Dest::Stack, line);
}

fn emit_str_replace_count_store(
    chunks: &mut [Chunk],
    current: usize,
    ref_slot: u16,
    count_slot: u16,
    line: u32,
) {
    use vybe_compiler::primitives::pointers::{
        CARRAY_BASE_KEY, CARRAY_IDX_KEY, CARRAY_KIND, CELL_KIND, REF_KIND_KEY,
    };

    let chunk = &mut chunks[current];
    let kind_slot = alloc_local(chunk);
    lget(chunk, ref_slot, line);
    emit_class_get_internal(chunk, REF_KIND_KEY, line);
    lset(chunk, kind_slot, line);

    lget(chunk, kind_slot, line);
    push_str(chunk, CELL_KIND, line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);

    lget(chunk, ref_slot, line);
    vybe_compiler::primitives::references::emit_cell_store(chunks, current, count_slot, line);
    let chunk = &mut chunks[current];

    chunk.emit_else(line);
    lget(chunk, kind_slot, line);
    push_str(chunk, CARRAY_KIND, line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);

    let base_slot = alloc_local(chunk);
    let idx_slot = alloc_local(chunk);
    lget(chunk, ref_slot, line);
    emit_class_get_internal(chunk, CARRAY_BASE_KEY, line);
    lset(chunk, base_slot, line);
    lget(chunk, ref_slot, line);
    emit_class_get_internal(chunk, CARRAY_IDX_KEY, line);
    lset(chunk, idx_slot, line);

    lget(chunk, base_slot, line);
    emit_class_get_internal(chunk, REF_KIND_KEY, line);
    push_str(chunk, CELL_KIND, line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);

    lget(chunk, base_slot, line);
    vybe_compiler::primitives::references::emit_cell_store(chunks, current, count_slot, line);
    let chunk = &mut chunks[current];

    chunk.emit_else(line);
    lget(chunk, base_slot, line);
    lget(chunk, idx_slot, line);
    lget(chunk, count_slot, line);
    vybe_compiler::primitives::collections::emit_set(chunks, current, line);
    let chunk = &mut chunks[current];
    chunk.emit_op(Op::DROP, line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_end(line);
}

pub fn emit_str_replace(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let subj_slot = alloc_local(chunk);
    let repl_slot = alloc_local(chunk);
    let srch_slot = alloc_local(chunk);
    let count_slot = alloc_local(chunk);
    let has_count = argc >= 4;
    let count_arg_slot = if has_count {
        let slot = alloc_local(chunk);
        lset(chunk, slot, line);
        Some(slot)
    } else {
        None
    };
    lset(chunk, subj_slot, line);
    lset(chunk, repl_slot, line);
    lset(chunk, srch_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    lset(chunk, count_slot, line);

    // Coerce subj to string — PHP's coercion, not ECMA's. This was a fourth
    // inline `"" + v`, which is why `str_replace("1","Z",false)` yielded `"0"`
    // where real php yields `""`.
    lget(chunk, subj_slot, line);
    coerce_to_str(chunk, line);
    lset(chunk, subj_slot, line);

    // is_array_search = Array.isArray(srch)
    let _ = chunk;
    chunks[current].emit_op_u16(Op::LOCAL_GET, srch_slot, line);
    call_import(chunks, current, "ecma:array", "isArray", 1, line);
    let chunk = &mut chunks[current];
    // ecma:array.isArray returns I32(0|1); structured if: array path if true, scalar path if false.
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line); // if is_array: array path; else: scalar path

    // ── Array path ──
    // is_array_repl = Array.isArray(repl)
    let is_array_repl_slot = alloc_local(chunk);
    push_const(chunk, Value::Bool(false), line);
    lset(chunk, is_array_repl_slot, line);
    let _ = chunk;
    chunks[current].emit_op_u16(Op::LOCAL_GET, repl_slot, line);
    call_import(chunks, current, "ecma:array", "isArray", 1, line);
    let chunk = &mut chunks[current];
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    push_const(chunk, Value::Bool(true), line);
    lset(chunk, is_array_repl_slot, line);
    chunk.emit_end(line);

    // for i in 0..srch.length: needle = srch[i]; rep = ...; subj = ecma:string.replaceAll(subj, needle, rep)
    let i_slot = alloc_local(chunk);
    let n_slot = alloc_local(chunk);
    let needle_slot = alloc_local(chunk);
    let rep_slot = alloc_local(chunk);
    push_const(chunk, Value::F64(0.0), line);
    lset(chunk, i_slot, line);
    lget(chunk, srch_slot, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);
    lset(chunk, n_slot, line);

    let _ = chunk;
    let loop_state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, i_slot, line);
    lget(chunk, n_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);
    let chunk = &mut chunks[current];

    // needle = (string) srch[i]
    lget(chunk, srch_slot, line);
    lget(chunk, i_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    coerce_to_str(chunk, line);
    lset(chunk, needle_slot, line);

    // if needle.length > 0: do replacement (else skip)
    lget(chunk, needle_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "length");
        chunk.emit_call(idx, 1, line);
    }
    push_const(chunk, Value::F64(0.0), line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_not(chunk, line);
    chunk.emit_if(line);

    // rep = is_array_repl ? (i < repl.length ? "" + repl[i] : "") : "" + repl
    lget(chunk, is_array_repl_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    // array repl path
    lget(chunk, i_slot, line);
    lget(chunk, repl_slot, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    lget(chunk, repl_slot, line);
    lget(chunk, i_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    coerce_to_str(chunk, line);
    chunk.emit_else(line);
    push_str(chunk, "", line);
    chunk.emit_end(line);
    chunk.emit_else(line);
    // scalar repl path
    lget(chunk, repl_slot, line);
    coerce_to_str(chunk, line);
    chunk.emit_end(line);
    lset(chunk, rep_slot, line);

    // count += substr_count(subj, needle) before subject is mutated.
    let _ = chunk;
    emit_str_replace_count_occurrences(chunks, current, subj_slot, needle_slot, count_slot, line);
    let chunk = &mut chunks[current];

    // subj = ecma:string.replaceAll(subj, needle, rep)
    lget(chunk, subj_slot, line);
    lget(chunk, needle_slot, line);
    lget(chunk, rep_slot, line);
    {
        let idx = chunk.add_import("ecma:string", "replaceAll");
        chunk.emit_call(idx, 3, line);
    }
    lset(chunk, subj_slot, line);

    chunk.emit_end(line); // end if nonempty_needle

    // i++
    lget(chunk, i_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, i_slot, line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, loop_state, line);
    let chunk = &mut chunks[current];

    lget(chunk, subj_slot, line);
    chunk.emit_else(line); // else branch of scalar_path if-value

    // ── Scalar path ──
    // ecma:string.replaceAll(subj, (string) srch, (string) repl), unless the
    // needle is empty. PHP treats an empty search as no replacements.
    let scalar_needle_slot = alloc_local(chunk);
    let scalar_rep_slot = alloc_local(chunk);
    lget(chunk, srch_slot, line);
    coerce_to_str(chunk, line);
    lset(chunk, scalar_needle_slot, line);
    lget(chunk, repl_slot, line);
    coerce_to_str(chunk, line);
    lset(chunk, scalar_rep_slot, line);

    lget(chunk, scalar_needle_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "length");
        chunk.emit_call(idx, 1, line);
    }
    push_const(chunk, Value::F64(0.0), line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_not(chunk, line);
    chunk.emit_if(line); // Resultless scalar nonempty body: exact native type repair.

    let _ = chunk;
    emit_str_replace_count_occurrences(
        chunks,
        current,
        subj_slot,
        scalar_needle_slot,
        count_slot,
        line,
    );
    let chunk = &mut chunks[current];

    lget(chunk, subj_slot, line);
    lget(chunk, scalar_needle_slot, line);
    lget(chunk, scalar_rep_slot, line);
    {
        let idx = chunk.add_import("ecma:string", "replaceAll");
        chunk.emit_call(idx, 3, line);
    }
    lset(chunk, subj_slot, line);
    chunk.emit_end(line);
    lget(chunk, subj_slot, line);

    chunk.emit_end(line); // end scalar_path if-value

    if let Some(count_arg_slot) = count_arg_slot {
        lset(chunk, subj_slot, line);
        let _ = chunk;
        emit_str_replace_count_store(chunks, current, count_arg_slot, count_slot, line);
        let chunk = &mut chunks[current];
        lget(chunk, subj_slot, line);
    }
}

pub fn emit_count_capture(
    chunks: &mut [Chunk],
    current: usize,
    subj: u16,
    needle: u16,
    count: u16,
    line: u32,
) {
    emit_str_replace_count_occurrences(chunks, current, subj, needle, count, line);
}
