//! PHP math helpers — Rust inline opcode emitters.
//!
//! Implements `min` / `max` (variadic + array form), `decbin` /
//! `decoct` / `dechex` (radix-string conversion), `base_convert`
//! (string↔string base conversion). Composes only WASM ops +
//! `ecma:math.fmod` (for `%` on floats); no PHP-specific host fns.

use vybe_bytecode::{Chunk, Value};
use vybe_bytecode::opcode::Op;
use std::sync::Arc;

fn alloc_local(chunk: &mut Chunk) -> u16 {
    let s = chunk.local_count;
    chunk.local_count = s + 1;
    s
}

fn push_const(chunk: &mut Chunk, val: Value, line: u32) {
    let idx = chunk.add_constant(val);
    chunk.emit_op_u16(Op::CONST, idx, line);
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
    emit_min_or_max(chunks, current, argc, /*want_lt=*/true, line);
}

/// PHP `max(...)`. Same shape as `min`.
pub fn emit_php_max(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_min_or_max(chunks, current, argc, /*want_lt=*/false, line);
}

fn emit_min_or_max(chunks: &mut [Chunk], current: usize, argc: u8, want_lt: bool, line: u32) {
    let chunk = &mut chunks[current];
    if argc == 0 {
        chunk.emit_op(Op::FALSE, line);
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
        // tmp = +arg[i]
        lget(chunk, base + i as u16, line);
        push_const(chunk, Value::F64(0.0), line);
        chunk.emit_op(Op::DYN_ADD, line); // numeric coerce via "0 + v"
        // best_n = +best
        lget(chunk, best_slot, line);
        push_const(chunk, Value::F64(0.0), line);
        chunk.emit_op(Op::DYN_ADD, line);
        chunk.emit_op(if want_lt { Op::DYN_LT } else { Op::DYN_GT }, line);
        let skip = chunk.emit_jump(Op::BR_IF_FALSE, line);
        lget(chunk, base + i as u16, line);
        lset(chunk, best_slot, line);
        chunk.patch_jump(skip);
    }
    lget(chunk, best_slot, line);
}

/// Fold `arr` to the min/max element. `arr_slot` holds the array.
/// Stack on entry: `[]` ; Stack on exit: `[result]`.
fn emit_reduce_array(chunk: &mut Chunk, arr_slot: u16, want_lt: bool, line: u32) {
    let len_slot = alloc_local(chunk);
    let i_slot = alloc_local(chunk);
    let best_slot = alloc_local(chunk);

    // len = arr.length
    lget(chunk, arr_slot, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);
    lset(chunk, len_slot, line);

    // if len === 0: push false and exit
    lget(chunk, len_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    chunk.emit_op(Op::DYN_EQ, line);
    let nonempty = chunk.emit_jump(Op::BR_IF_FALSE, line);
    chunk.emit_op(Op::FALSE, line);
    let done_empty = chunk.emit_jump(Op::BR, line);
    chunk.patch_jump(nonempty);

    // best = arr[0]
    lget(chunk, arr_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    chunk.emit_op(Op::ARRAY_GET, line);
    lset(chunk, best_slot, line);

    // i = 1
    push_const(chunk, Value::F64(1.0), line);
    lset(chunk, i_slot, line);

    let loop_top = chunk.current_offset();
    lget(chunk, i_slot, line);
    lget(chunk, len_slot, line);
    chunk.emit_op(Op::DYN_LT, line);
    let exit = chunk.emit_jump(Op::BR_IF_FALSE, line);

    // tmp = +arr[i]
    lget(chunk, arr_slot, line);
    lget(chunk, i_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    push_const(chunk, Value::F64(0.0), line);
    chunk.emit_op(Op::DYN_ADD, line);
    // best_n = +best
    lget(chunk, best_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    chunk.emit_op(Op::DYN_ADD, line);
    chunk.emit_op(if want_lt { Op::DYN_LT } else { Op::DYN_GT }, line);
    let skip = chunk.emit_jump(Op::BR_IF_FALSE, line);
    lget(chunk, arr_slot, line);
    lget(chunk, i_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    lset(chunk, best_slot, line);
    chunk.patch_jump(skip);

    // i++
    lget(chunk, i_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, i_slot, line);
    chunk.emit_loop(loop_top, line);
    chunk.patch_jump(exit);

    lget(chunk, best_slot, line);
    chunk.patch_jump(done_empty);
}

pub fn emit_php_decbin(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    emit_dec_to_radix(chunks, current, 2, /*hex_digits=*/false, line);
}

pub fn emit_php_decoct(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    emit_dec_to_radix(chunks, current, 8, /*hex_digits=*/false, line);
}

pub fn emit_php_dechex(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    emit_dec_to_radix(chunks, current, 16, /*hex_digits=*/true, line);
}

/// Stack on entry: `[n]` ; Stack on exit: `[radix-string]`.
fn emit_dec_to_radix(chunks: &mut [Chunk], current: usize, radix: i32, hex_digits: bool, line: u32) {
    let chunk = &mut chunks[current];
    let n_slot = alloc_local(chunk);
    let out_slot = alloc_local(chunk);
    let digit_slot = alloc_local(chunk);

    // n = +arg
    push_const(chunk, Value::F64(0.0), line);
    chunk.emit_op(Op::DYN_ADD, line);
    lset(chunk, n_slot, line);

    // if n === 0: push "0" and BR to end
    lget(chunk, n_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    chunk.emit_op(Op::DYN_EQ, line);
    let nonzero = chunk.emit_jump(Op::BR_IF_FALSE, line);
    push_str(chunk, "0", line);
    let done_zero = chunk.emit_jump(Op::BR, line);
    chunk.patch_jump(nonzero);

    // out = ""
    push_str(chunk, "", line);
    lset(chunk, out_slot, line);

    let fmod = chunks[0].add_import("ecma:math", "fmod");
    let chunk = &mut chunks[current];

    // while n > 0
    let loop_top = chunk.current_offset();
    lget(chunk, n_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    chunk.emit_op(Op::DYN_GT, line);
    let exit = chunk.emit_jump(Op::BR_IF_FALSE, line);

    // digit = n % radix
    lget(chunk, n_slot, line);
    push_const(chunk, Value::F64(radix as f64), line);
    chunk.emit_op_u16(Op::CALL_IMPORT, fmod, line);
    chunk.emit(2, line);
    lset(chunk, digit_slot, line);

    // out = (table.charAt(digit) | String(digit)) + out
    if hex_digits {
        push_str(chunk, "0123456789abcdef", line);
        lget(chunk, digit_slot, line);
        chunk.emit_op(Op::STR_CHAR_AT, line);
    } else {
        push_str(chunk, "", line);
        lget(chunk, digit_slot, line);
        chunk.emit_op(Op::DYN_ADD, line);
    }
    lget(chunk, out_slot, line);
    chunk.emit_op(Op::DYN_ADD, line);
    lset(chunk, out_slot, line);

    // n = floor(n / radix)
    lget(chunk, n_slot, line);
    push_const(chunk, Value::F64(radix as f64), line);
    chunk.emit_op(Op::F64_DIV, line);
    chunk.emit_op(Op::F64_FLOOR, line);
    lset(chunk, n_slot, line);

    chunk.emit_loop(loop_top, line);
    chunk.patch_jump(exit);

    lget(chunk, out_slot, line);
    chunk.patch_jump(done_zero);
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
    chunk.emit_op(Op::DYN_ADD, line);
    chunk.emit_op(Op::STR_TO_LOWER, line);
    lset(chunk, s_slot, line);

    let table_str = "0123456789abcdefghijklmnopqrstuvwxyz";

    // n = 0 ; len = s.length ; i = 0
    push_const(chunk, Value::F64(0.0), line);
    lset(chunk, n_slot, line);
    lget(chunk, s_slot, line);
    chunk.emit_op(Op::STR_LENGTH, line);
    lset(chunk, len_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    lset(chunk, i_slot, line);

    let loop_top = chunk.current_offset();
    lget(chunk, i_slot, line);
    lget(chunk, len_slot, line);
    chunk.emit_op(Op::DYN_LT, line);
    let exit = chunk.emit_jump(Op::BR_IF_FALSE, line);

    // d = table.indexOf(s.charAt(i))
    push_str(chunk, table_str, line);
    lget(chunk, s_slot, line);
    lget(chunk, i_slot, line);
    chunk.emit_op(Op::STR_CHAR_AT, line);
    chunk.emit_op(Op::STR_INDEX_OF, line);
    lset(chunk, d_slot, line);

    // if d >= 0 && d < from: n = n*from + d
    lget(chunk, d_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    chunk.emit_op(Op::DYN_LT, line);
    chunk.emit_op(Op::DYN_NOT, line);
    let skip = chunk.emit_jump(Op::BR_IF_FALSE, line);
    lget(chunk, d_slot, line);
    lget(chunk, from_slot, line);
    chunk.emit_op(Op::DYN_LT, line);
    let skip2 = chunk.emit_jump(Op::BR_IF_FALSE, line);
    lget(chunk, n_slot, line);
    lget(chunk, from_slot, line);
    chunk.emit_op(Op::F64_MUL, line);
    lget(chunk, d_slot, line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, n_slot, line);
    chunk.patch_jump(skip2);
    chunk.patch_jump(skip);

    // i++
    lget(chunk, i_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, i_slot, line);
    chunk.emit_loop(loop_top, line);
    chunk.patch_jump(exit);

    // if n === 0: push "0" and BR to end
    lget(chunk, n_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    chunk.emit_op(Op::DYN_EQ, line);
    let nonzero = chunk.emit_jump(Op::BR_IF_FALSE, line);
    push_str(chunk, "0", line);
    let done_zero = chunk.emit_jump(Op::BR, line);
    chunk.patch_jump(nonzero);

    // out = ""
    push_str(chunk, "", line);
    lset(chunk, out_slot, line);

    let fmod = chunks[0].add_import("ecma:math", "fmod");
    let chunk = &mut chunks[current];

    // while n > 0: out = table.charAt(n % to) + out; n = floor(n / to)
    let loop2_top = chunk.current_offset();
    lget(chunk, n_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    chunk.emit_op(Op::DYN_GT, line);
    let exit2 = chunk.emit_jump(Op::BR_IF_FALSE, line);

    // digit_idx = n % to
    lget(chunk, n_slot, line);
    lget(chunk, to_slot, line);
    chunk.emit_op_u16(Op::CALL_IMPORT, fmod, line);
    chunk.emit(2, line);
    let digit_slot = alloc_local(chunk);
    lset(chunk, digit_slot, line);

    push_str(chunk, table_str, line);
    lget(chunk, digit_slot, line);
    chunk.emit_op(Op::STR_CHAR_AT, line);
    lget(chunk, out_slot, line);
    chunk.emit_op(Op::DYN_ADD, line);
    lset(chunk, out_slot, line);

    // n = floor(n / to)
    lget(chunk, n_slot, line);
    lget(chunk, to_slot, line);
    chunk.emit_op(Op::F64_DIV, line);
    chunk.emit_op(Op::F64_FLOOR, line);
    lset(chunk, n_slot, line);

    chunk.emit_loop(loop2_top, line);
    chunk.patch_jump(exit2);

    lget(chunk, out_slot, line);
    chunk.patch_jump(done_zero);
}
