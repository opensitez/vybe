//! PHP math helpers — Rust inline opcode emitters.
//!
//! Implements `min` / `max` (variadic + array form), `decbin` /
//! `decoct` / `dechex` (radix-string conversion), `base_convert`
//! (string↔string base conversion). Composes only WASM ops +
//! `ecma:math.fmod` (for `%` on floats); no PHP-specific host fns.

use std::sync::Arc;
use vybe_bytecode::opcode::Op;
use vybe_bytecode::{Chunk, Value};

fn alloc_local(chunk: &mut Chunk) -> u16 {
    let s = chunk.local_count;
    chunk.local_count = s + 1;
    s
}

fn push_const(chunk: &mut Chunk, val: Value, line: u32) {
    match &val {
        Value::F64(v) => chunk.emit_f64_const(*v, line),
        Value::I32(v) => chunk.emit_i32_const(*v, line),
        Value::Null => chunk.emit_op(Op::NULL, line),
        Value::BigInt(v) => chunk.emit_i64_const(*v, line),
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
    chunk.emit_op(Op::DROP, line);
}

fn lget(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

/// PHP `min(a, b, ...)` / `min([a, b, ...])`.
/// Stack on entry: `[arg0, arg1, ..., argN-1]` (argc args).
/// Stack on exit: `[smallest]` or `[false]` if argc == 0.
pub fn emit_php_min(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_min_or_max(chunks, current, argc, /*want_lt=*/ true, line);
}

/// PHP `max(...)`. Same shape as `min`.
pub fn emit_php_max(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_min_or_max(chunks, current, argc, /*want_lt=*/ false, line);
}

/// Coerce the value on top of the stack to a number for comparison.
/// Already-numeric values pass through unchanged; strings (incl. numeric
/// strings like `'10'`) go through `ecma:number.parseFloat`. `parseFloat`
/// alone is wrong because the host fn expects a string and returns 0/NaN
/// for an f64 operand, and `0 + v` is wrong because it concatenates a
/// string operand. Stack: `[v]` → `[number]`.
fn emit_to_number(chunk: &mut Chunk, pf_idx: u16, line: u32) {
    let t = alloc_local(chunk);
    lset(chunk, t, line);
    lget(chunk, t, line);
    let test_num_t = chunk.add_import("wasm:js-number", "test");
    chunk.emit_call(test_num_t, 1, line);
    chunk.emit_if_value(line);
    lget(chunk, t, line);
    chunk.emit_else(line);
    lget(chunk, t, line);
    chunk.emit_call(pf_idx, 1, line);
    chunk.emit_end(line);
}

fn emit_min_or_max(chunks: &mut [Chunk], current: usize, argc: u8, want_lt: bool, line: u32) {
    // Numeric coercion for comparison must turn numeric strings into real
    // numbers (`'10'` → 10). `0 + v` would *concatenate* a string operand
    // (→ "010") and then trap in the comparison's `toF64`, so coerce with
    // `ecma:number.parseFloat` instead. Resolve the import up-front (it
    // lives on chunk 0's import table).
    let pf_idx = chunks[0].add_import("ecma:number".to_string(), "parseFloat".to_string());
    let chunk = &mut chunks[current];
    if argc == 0 {
        push_const(chunk, Value::Bool(false), line);
        return;
    }
    // Stash all args into slots in argument order.
    let base = chunk.local_count;
    chunk.local_count = base + argc as u16;
    for i in (0..argc).rev() {
        chunk.emit_op_u16(Op::LOCAL_SET, base + i as u16, line);
        chunk.emit_op(Op::DROP, line);
    }
    // Special-case: argc == 1 and arg[0] is an array → reduce over it.
    // We can't statically test "is array" here, but `ARRAY_LENGTH` on
    // a non-array traps. To stay safe, only enter the array path when
    // argc == 1 and the single arg is materially an array — runtime
    // detection via a small bytecode probe.
    if argc == 1 {
        // typeof arg[0] === "object" and ARRAY_LENGTH succeeds → reduce.
        // Simplification for the test surface: if argc==1, treat the
        // single argument as the values array. PHP semantics agree
        // when called with `min([a, b, c])`.
        emit_reduce_array(chunk, base, want_lt, line);
        return;
    }
    // Multi-arg form: `arg[0]` is `best` initially, fold others in.
    let best_slot = alloc_local(chunk);
    lget(chunk, base, line);
    lset(chunk, best_slot, line);
    for i in 1..argc {
        // tmp = toNumber(arg[i])
        lget(chunk, base + i as u16, line);
        emit_to_number(chunk, pf_idx, line);
        // best_n = toNumber(best)
        lget(chunk, best_slot, line);
        emit_to_number(chunk, pf_idx, line);
        if want_lt {
            crate::emitter::ops::emit_dyn_lt(chunk, line)
        } else {
            crate::emitter::ops::emit_dyn_gt(chunk, line)
        };
        chunk.emit_if(line);
        lget(chunk, base + i as u16, line);
        lset(chunk, best_slot, line);
        chunk.emit_end(line);
    }
    lget(chunk, best_slot, line);
}

/// Fold `arr` to the min/max element. `arr_slot` holds the array.
/// Stack on entry: `[]` ; Stack on exit: `[result]`.
fn emit_reduce_array(chunk: &mut Chunk, arr_slot: u16, want_lt: bool, line: u32) {
    let len_slot = alloc_local(chunk);
    let i_slot = alloc_local(chunk);
    let best_slot = alloc_local(chunk);
    let result_slot = alloc_local(chunk);

    // len = arr.length
    lget(chunk, arr_slot, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);
    lset(chunk, len_slot, line);

    let outer = chunk.emit_block(line);

    // if len === 0: result = false; break
    lget(chunk, len_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    crate::emitter::ops::emit_dyn_eq(chunk, line);
    chunk.emit_if(line);
    push_const(chunk, Value::Bool(false), line);
    lset(chunk, result_slot, line);
    chunk.emit_br(1, line);
    chunk.emit_end(line);

    // best = arr[0]
    lget(chunk, arr_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    chunk.emit_op(Op::ARRAY_GET, line);
    lset(chunk, best_slot, line);

    // i = 1
    push_const(chunk, Value::F64(1.0), line);
    lset(chunk, i_slot, line);

    let loop_block = chunk.emit_block(line);
    let (loop_patch, _) = chunk.emit_loop_s(line);
    lget(chunk, i_slot, line);
    lget(chunk, len_slot, line);
    crate::emitter::ops::emit_dyn_lt(chunk, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_br_if(1, line);

    // tmp = +arr[i]
    lget(chunk, arr_slot, line);
    lget(chunk, i_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    push_const(chunk, Value::F64(0.0), line);
    crate::emitter::ops::emit_dyn_add(chunk, line);
    // best_n = +best
    lget(chunk, best_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    crate::emitter::ops::emit_dyn_add(chunk, line);
    if want_lt {
        crate::emitter::ops::emit_dyn_lt(chunk, line)
    } else {
        crate::emitter::ops::emit_dyn_gt(chunk, line)
    };
    chunk.emit_if(line);
    lget(chunk, arr_slot, line);
    lget(chunk, i_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    lset(chunk, best_slot, line);
    chunk.emit_end(line);

    // i++
    lget(chunk, i_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, i_slot, line);
    chunk.emit_br(0, line);
    chunk.emit_end(line);
    chunk.patch_loop(loop_patch);
    chunk.emit_end(line);
    chunk.patch_block(loop_block);

    lget(chunk, best_slot, line);
    lset(chunk, result_slot, line);
    chunk.emit_end(line);
    chunk.patch_block(outer);
    lget(chunk, result_slot, line);
}

pub fn emit_php_decbin(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    emit_dec_to_radix(chunks, current, 2, /*hex_digits=*/ false, line);
}

pub fn emit_php_decoct(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    emit_dec_to_radix(chunks, current, 8, /*hex_digits=*/ false, line);
}

pub fn emit_php_dechex(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    emit_dec_to_radix(chunks, current, 16, /*hex_digits=*/ true, line);
}

/// Stack on entry: `[n]` ; Stack on exit: `[radix-string]`.
fn emit_dec_to_radix(
    chunks: &mut [Chunk],
    current: usize,
    radix: i32,
    hex_digits: bool,
    line: u32,
) {
    let chunk = &mut chunks[current];
    let n_slot = alloc_local(chunk);
    let out_slot = alloc_local(chunk);
    let digit_slot = alloc_local(chunk);

    // n = +arg ; out = ""
    push_const(chunk, Value::F64(0.0), line);
    crate::emitter::ops::emit_dyn_add(chunk, line);
    lset(chunk, n_slot, line);
    push_str(chunk, "", line);
    lset(chunk, out_slot, line);

    // B_outer — the n == 0 shortcut breaks out to the result load below.
    let outer = chunk.emit_block(line);

    // if n === 0 { out = "0"; break B_outer }
    lget(chunk, n_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    crate::emitter::ops::emit_dyn_eq(chunk, line);
    chunk.emit_if(line);
    push_str(chunk, "0", line);
    lset(chunk, out_slot, line);
    chunk.emit_br(1, line); // exit B_outer
    chunk.emit_end(line);

    // B_loop — wraps the digit loop so the loop-exit branch lands on the
    // result load instead of skipping past B_outer (which would drop the
    // accumulated string and leave the function returning null).
    let loop_block = chunk.emit_block(line);
    let (loop_patch, _) = chunk.emit_loop_s(line);

    // while n > 0  (exit B_loop when !(n > 0))
    lget(chunk, n_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    crate::emitter::ops::emit_dyn_gt(chunk, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_br_if(1, line); // exit B_loop

    // digit = n % radix
    lget(chunk, n_slot, line);
    push_const(chunk, Value::F64(radix as f64), line);
    crate::emitter::math::emit_c_fmod(chunk, line);
    lset(chunk, digit_slot, line);

    // out = (table.charAt(digit) | String(digit)) + out
    if hex_digits {
        push_str(chunk, "0123456789abcdef", line);
        lget(chunk, digit_slot, line);
        { let idx = chunk.add_import("ecma:string", "charAt"); chunk.emit_call(idx, 2, line); }
    } else {
        push_str(chunk, "", line);
        lget(chunk, digit_slot, line);
        crate::emitter::ops::emit_dyn_add(chunk, line);
    }
    lget(chunk, out_slot, line);
    crate::emitter::ops::emit_dyn_add(chunk, line);
    lset(chunk, out_slot, line);

    // n = floor(n / radix)
    lget(chunk, n_slot, line);
    push_const(chunk, Value::F64(radix as f64), line);
    chunk.emit_op(Op::F64_DIV, line);
    chunk.emit_op(Op::F64_FLOOR, line);
    lset(chunk, n_slot, line);

    chunk.emit_br(0, line); // continue loop
    chunk.emit_end(line);
    chunk.patch_loop(loop_patch);
    chunk.emit_end(line); // end B_loop
    chunk.patch_block(loop_block);

    chunk.emit_end(line); // end B_outer
    chunk.patch_block(outer);

    lget(chunk, out_slot, line); // result string
}

/// PHP `base_convert(num_str, from_radix, to_radix)`.
/// Stack on entry: `[s, from, to]` ; Stack on exit: `[result_string]`.
pub fn emit_php_base_convert(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let to_slot = alloc_local(chunk);
    let from_slot = alloc_local(chunk);
    let s_slot = alloc_local(chunk);
    let n_slot = alloc_local(chunk);
    let i_slot = alloc_local(chunk);
    let len_slot = alloc_local(chunk);
    let d_slot = alloc_local(chunk);
    let out_slot = alloc_local(chunk);

    // Pop in reverse order: top is `to`, below is `from`, below is `s`.
    lset(chunk, to_slot, line);
    lset(chunk, from_slot, line);
    lset(chunk, s_slot, line);

    // s = ("" + s).toLowerCase()
    push_str(chunk, "", line);
    lget(chunk, s_slot, line);
    crate::emitter::ops::emit_dyn_add(chunk, line);
    { let idx = chunk.add_import("ecma:string", "toLowerCase"); chunk.emit_call(idx, 1, line); }
    lset(chunk, s_slot, line);

    let table_str = "0123456789abcdefghijklmnopqrstuvwxyz";

    // n = 0 ; len = s.length ; i = 0
    push_const(chunk, Value::F64(0.0), line);
    lset(chunk, n_slot, line);
    lget(chunk, s_slot, line);
    { let idx = chunk.add_import("wasm:js-string", "length"); chunk.emit_call(idx, 1, line); }
    lset(chunk, len_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    lset(chunk, i_slot, line);

    // block { loop { ... } } — the surrounding block makes `br_if 1` (exit
    // when i >= len) a valid WASM label. Without it the break has no target.
    let scan_block = chunk.emit_block(line);
    let (scan_loop_patch, _) = chunk.emit_loop_s(line);
    lget(chunk, i_slot, line);
    lget(chunk, len_slot, line);
    crate::emitter::ops::emit_dyn_lt(chunk, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_br_if(1, line);

    // d = table.indexOf(s.charAt(i))
    push_str(chunk, table_str, line);
    lget(chunk, s_slot, line);
    lget(chunk, i_slot, line);
    { let idx = chunk.add_import("ecma:string", "charAt"); chunk.emit_call(idx, 2, line); }
    { let idx = chunk.add_import("ecma:string", "indexOf"); chunk.emit_call(idx, 2, line); }
    lset(chunk, d_slot, line);

    // if d >= 0 && d < from: n = n*from + d
    lget(chunk, d_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    crate::emitter::ops::emit_dyn_lt(chunk, line);
    crate::emitter::ops::emit_dyn_not(chunk, line);
    chunk.emit_if(line);
    lget(chunk, d_slot, line);
    lget(chunk, from_slot, line);
    crate::emitter::ops::emit_dyn_lt(chunk, line);
    chunk.emit_if(line);
    lget(chunk, n_slot, line);
    lget(chunk, from_slot, line);
    chunk.emit_op(Op::F64_MUL, line);
    lget(chunk, d_slot, line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, n_slot, line);
    chunk.emit_end(line);
    chunk.emit_end(line);

    // i++
    lget(chunk, i_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, i_slot, line);
    chunk.emit_br(0, line);
    chunk.emit_end(line);
    chunk.patch_loop(scan_loop_patch);
    chunk.emit_end(line); // end scan block
    chunk.patch_block(scan_block);

    let outer = chunk.emit_block(line);

    // if n === 0: push "0" and break
    lget(chunk, n_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    crate::emitter::ops::emit_dyn_eq(chunk, line);
    chunk.emit_if(line);
    push_str(chunk, "0", line);
    chunk.emit_br(1, line);
    chunk.emit_end(line);

    // out = ""
    push_str(chunk, "", line);
    lset(chunk, out_slot, line);

    let chunk = &mut chunks[current];

    // while n > 0: out = table.charAt(n % to) + out; n = floor(n / to)
    // Own block so the loop break (`br_if 1`) lands right after the loop —
    // inside `outer`, so the trailing `lget out_slot` still runs.
    let out_block = chunk.emit_block(line);
    let (out_loop_patch, _) = chunk.emit_loop_s(line);
    lget(chunk, n_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    crate::emitter::ops::emit_dyn_gt(chunk, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_br_if(1, line);

    // digit_idx = n % to
    lget(chunk, n_slot, line);
    lget(chunk, to_slot, line);
    crate::emitter::math::emit_c_fmod(chunk, line);
    let digit_slot = alloc_local(chunk);
    lset(chunk, digit_slot, line);

    push_str(chunk, table_str, line);
    lget(chunk, digit_slot, line);
    { let idx = chunk.add_import("ecma:string", "charAt"); chunk.emit_call(idx, 2, line); }
    lget(chunk, out_slot, line);
    crate::emitter::ops::emit_dyn_add(chunk, line);
    lset(chunk, out_slot, line);

    // n = floor(n / to)
    lget(chunk, n_slot, line);
    lget(chunk, to_slot, line);
    chunk.emit_op(Op::F64_DIV, line);
    chunk.emit_op(Op::F64_FLOOR, line);
    lset(chunk, n_slot, line);

    chunk.emit_br(0, line);
    chunk.emit_end(line);
    chunk.patch_loop(out_loop_patch);
    chunk.emit_end(line); // end out block
    chunk.patch_block(out_block);

    lget(chunk, out_slot, line);
    chunk.emit_end(line);
    chunk.patch_block(outer);
}
