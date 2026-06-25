//! .NET `String.Format` / VB `Format` composite-format adapter — bytecode-only.
//!
//! `String.Format("Hello {0}, age {1}", name, age)` is .NET-style composite
//! formatting: the format string contains `{N}` placeholders that index into
//! the trailing args. Vybe lowers each call at compile time through this
//! adapter — pure inline bytecode using `STR_LENGTH` / `STR_SUBSTRING` /
//! `STR_CHAR_CODE_AT` / `STR_CONCAT` primitives. No host fns.
//!
//! Supported placeholder grammar (.NET §Composite formatting):
//!   `{{` / `}}`     literal `{` / `}`
//!   `{N}`           args[N] formatted with default `ToString()`
//!   `{N,W}`         padded to width W (right-align if W>0, left if W<0)
//!   `{N:fmt}` / `{N,W:fmt}` — format spec and alignment are routed through
//!                              the shared .NET numeric formatter.
//!
//! Call shape: at the call site, stack on entry is `[fmt, arg0, arg1, ..., argN-2]`
//! (so `argc` is the number of args including the format string). The adapter
//! packs `arg0..` into an array local, then walks the format string emitting
//! literal chars or `String(args[idx])` substitutions.

use crate::emitter::instructions::{core_wasm, host};
use std::sync::Arc;
use vybe_bytecode::opcode::Op;
use vybe_bytecode::{Chunk, Value};

fn push_const(chunk: &mut Chunk, val: Value, line: u32) {
    match &val {
        Value::String(s) => chunk.emit_string_const(s, line),
        Value::F64(f) => chunk.emit_f64_const(*f, line),
        Value::I32(i) => chunk.emit_i32_const(*i, line),
        _ => panic!("push_const: no WASM-compliant encoding for {:?}", val),
    }
}

fn emit_dotnet_format_value_call(
    chunk: &mut Chunk,
    args_slot: u16,
    idx_slot: u16,
    format_slot: u16,
    width_slot: u16,
    line: u32,
) {
    let helper = chunk.add_constant(Value::String(Arc::from("__vybe_dotnet_numeric_format")));
    chunk.emit_op_u16(Op::GLOBAL_GET, helper, line);
    chunk.emit_op_u16(Op::LOCAL_GET, args_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    chunk.emit_op_u16(Op::LOCAL_GET, format_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, width_slot, line);
    chunk.emit_op_u8(Op::CALL_REF, 3, line);
}

/// Emit `String.Format(fmt, ...args)` at the call site.
///
/// Stack on entry: `[fmt, arg_0, arg_1, ..., arg_{n-1}]` where
/// `argc == n + 1` (format string + n placeholder values).
/// Stack on exit: `[result_string]`.
pub fn emit_string_format(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    if argc == 0 {
        // Defensive: no format string was pushed. Emit empty string.
        push_const(chunk, Value::String(Arc::from("")), line);
        return;
    }

    // Stash trailing args in an array local so the format walker can
    // index into them by `{N}` placeholder.
    let n = (argc as u16) - 1;
    let args_slot = chunk.local_count;
    chunk.local_count = args_slot + 1;

    // Build the args array from the top `n` stack entries.
    // `ARRAY_NEW_FIXED n` pops the top n values (in stack order) and
    // builds an array — preserves order, so args[0] is the first
    // placeholder value as written in source.
    chunk.emit_op_u16(Op::ARRAY_NEW_FIXED, n, line);

    // Save args array to a local; stack now `[fmt]`.
    chunk.emit_op_u16(Op::LOCAL_SET, args_slot, line);

    // Now emit the runtime walker:
    //   fmt_slot   = current local
    //   i          = next local
    //   len        = next + 1
    //   out        = next + 2
    let fmt_slot = chunk.local_count;
    let i_slot = fmt_slot + 1;
    let len_slot = fmt_slot + 2;
    let out_slot = fmt_slot + 3;
    chunk.local_count = fmt_slot + 4;

    // fmt_slot = pop fmt
    chunk.emit_op_u16(Op::LOCAL_SET, fmt_slot, line);

    // i = 0
    core_wasm::i32_const(chunk, line, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, i_slot, line);

    // len = STR_LENGTH(fmt)
    chunk.emit_op_u16(Op::LOCAL_GET, fmt_slot, line);
    host::emit(chunk, "wasm:js-string", "length", 1, line);
    chunk.emit_op_u16(Op::LOCAL_SET, len_slot, line);

    // out = ""
    push_const(chunk, Value::String(Arc::from("")), line);
    chunk.emit_op_u16(Op::LOCAL_SET, out_slot, line);

    // Constants reused inside the loop.

    let outer_block = chunk.emit_block(line);
    let (loop_p, _) = chunk.emit_loop_s(line);

    // if i >= len: break
    chunk.emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, len_slot, line);
    crate::emitter::ops::emit_dyn_lt(chunk, line);
    crate::emitter::ops::emit_dyn_not(chunk, line);
    chunk.emit_br_if(1, line);

    // ch_code = fmt.charCodeAt(i)
    chunk.emit_op_u16(Op::LOCAL_GET, fmt_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, i_slot, line);
    host::emit(chunk, "wasm:js-string", "charCodeAt", 2, line);

    // Branch: open_brace? close_brace? literal?
    // Using DUP + compare cascade; SELECT not used because we need different actions per branch.
    // Implementation strategy: nested blocks for each "case" tag.

    // Save char code to a temp slot for repeated comparisons.
    let ch_slot = chunk.local_count;
    chunk.local_count = ch_slot + 1;
    chunk.emit_op_u16(Op::LOCAL_SET, ch_slot, line);

    // -- '{' branch --
    let open_block = chunk.emit_block(line);
    chunk.emit_op_u16(Op::LOCAL_GET, ch_slot, line);
    chunk.emit_i32_const(b'{' as i32, line);
    crate::emitter::ops::emit_dyn_eq(chunk, line);
    crate::emitter::ops::emit_dyn_not(chunk, line);
    chunk.emit_br_if(0, line); // skip past open-brace handler if not '{'

    emit_handle_open_brace(
        chunk,
        fmt_slot,
        i_slot,
        len_slot,
        out_slot,
        args_slot,
        line,
    );
    // br(1) = continue loop (depth 0 = open_block, 1 = loop, 2 = outer)
    chunk.emit_br(1, line);
    chunk.emit_end(line);
    chunk.patch_block(open_block);

    // -- '}' branch (escape `}}` → `}`) --
    let close_block = chunk.emit_block(line);
    chunk.emit_op_u16(Op::LOCAL_GET, ch_slot, line);
    chunk.emit_i32_const(b'}' as i32, line);
    crate::emitter::ops::emit_dyn_eq(chunk, line);
    crate::emitter::ops::emit_dyn_not(chunk, line);
    chunk.emit_br_if(0, line);

    // peek next char; if also '}' append literal '}' and skip 2, else skip 1.
    emit_handle_close_brace(
        chunk,
        fmt_slot,
        i_slot,
        len_slot,
        out_slot,
        line,
    );
    chunk.emit_br(1, line);
    chunk.emit_end(line);
    chunk.patch_block(close_block);

    // -- literal char path: out += fmt.substring(i, i+1) --
    chunk.emit_op_u16(Op::LOCAL_GET, out_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, fmt_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, i_slot, line); // start
    chunk.emit_op_u16(Op::LOCAL_GET, i_slot, line); // end (about to add 1)
    core_wasm::i32_const(chunk, line, 1);
    chunk.emit_op(Op::I32_ADD, line);
    host::emit(chunk, "wasm:js-string", "substring", 3, line);
    crate::emitter::ops::emit_dyn_add(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_SET, out_slot, line);

    // i = i + 1
    chunk.emit_op_u16(Op::LOCAL_GET, i_slot, line);
    core_wasm::i32_const(chunk, line, 1);
    chunk.emit_op(Op::I32_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_SET, i_slot, line);

    chunk.emit_br(0, line);
    chunk.emit_end(line);
    chunk.patch_loop(loop_p);
    chunk.emit_end(line);
    chunk.patch_block(outer_block);

    // Push result
    chunk.emit_op_u16(Op::LOCAL_GET, out_slot, line);
}

/// Handle `{` at fmt[i]: either the `{{` escape or a `{N[,W][:fmt]}` placeholder.
/// On entry `i` points at `{`. Advances `i` past the consumed segment and
/// appends the formatted result to `out`.
fn emit_handle_open_brace(
    chunk: &mut Chunk,
    fmt_slot: u16,
    i_slot: u16,
    len_slot: u16,
    out_slot: u16,
    args_slot: u16,
    line: u32,
) {
    // Check if next char is also '{' → escape
    // peek = (i+1 < len) && fmt.charCodeAt(i+1) == '{'
    let escape_block = chunk.emit_block(line);
    chunk.emit_op_u16(Op::LOCAL_GET, i_slot, line);
    core_wasm::i32_const(chunk, line, 1);
    chunk.emit_op(Op::I32_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_GET, len_slot, line);
    crate::emitter::ops::emit_dyn_lt(chunk, line);
    crate::emitter::ops::emit_dyn_not(chunk, line);
    chunk.emit_br_if(0, line); // i+1 >= len → not escape

    chunk.emit_op_u16(Op::LOCAL_GET, fmt_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, i_slot, line);
    core_wasm::i32_const(chunk, line, 1);
    chunk.emit_op(Op::I32_ADD, line);
    host::emit(chunk, "wasm:js-string", "charCodeAt", 2, line);
    chunk.emit_i32_const(b'{' as i32, line);
    crate::emitter::ops::emit_dyn_eq(chunk, line);
    crate::emitter::ops::emit_dyn_not(chunk, line);
    chunk.emit_br_if(0, line); // not '{{' → not escape

    // It IS `{{`: append '{' to out, advance i by 2.
    chunk.emit_op_u16(Op::LOCAL_GET, out_slot, line);
    push_const(chunk, Value::String(Arc::from("{")), line);
    crate::emitter::ops::emit_dyn_add(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_SET, out_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunk.emit_i32_const(2, line);
    chunk.emit_op(Op::I32_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_SET, i_slot, line);
    // br depth 2 = continue outer loop. Depths from inside escape_block:
    // 0=escape_block, 1=open_block, 2=loop, 3=outer_block.
    chunk.emit_br(2, line);
    chunk.emit_end(line);
    chunk.patch_block(escape_block);

    // Not an escape: parse `{N[,W][:fmt]}` placeholder.
    // Find closing '}' starting from i+1.
    let end_slot = chunk.local_count;
    chunk.local_count = end_slot + 1;
    chunk.emit_op_u16(Op::LOCAL_GET, i_slot, line);
    core_wasm::i32_const(chunk, line, 1);
    chunk.emit_op(Op::I32_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_SET, end_slot, line);

    let scan_block = chunk.emit_block(line);
    let (scan_loop, _) = chunk.emit_loop_s(line);
    chunk.emit_op_u16(Op::LOCAL_GET, end_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, len_slot, line);
    crate::emitter::ops::emit_dyn_lt(chunk, line);
    crate::emitter::ops::emit_dyn_not(chunk, line);
    chunk.emit_br_if(1, line);

    // if fmt[end] == '}': break
    chunk.emit_op_u16(Op::LOCAL_GET, fmt_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, end_slot, line);
    host::emit(chunk, "wasm:js-string", "charCodeAt", 2, line);
    chunk.emit_i32_const(b'}' as i32, line);
    crate::emitter::ops::emit_dyn_eq(chunk, line);
    chunk.emit_br_if(1, line);

    // end += 1
    chunk.emit_op_u16(Op::LOCAL_GET, end_slot, line);
    core_wasm::i32_const(chunk, line, 1);
    chunk.emit_op(Op::I32_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_SET, end_slot, line);
    chunk.emit_br(0, line);
    chunk.emit_end(line);
    chunk.patch_loop(scan_loop);
    chunk.emit_end(line);
    chunk.patch_block(scan_block);

    // inner = fmt.substring(i+1, end) — the placeholder body, e.g. "0" or "0,5" or "0:N2"
    let inner_slot = chunk.local_count;
    chunk.local_count = inner_slot + 1;
    chunk.emit_op_u16(Op::LOCAL_GET, fmt_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, i_slot, line);
    core_wasm::i32_const(chunk, line, 1);
    chunk.emit_op(Op::I32_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_GET, end_slot, line);
    host::emit(chunk, "wasm:js-string", "substring", 3, line);
    chunk.emit_op_u16(Op::LOCAL_SET, inner_slot, line);

    let inner_len_slot = chunk.local_count;
    chunk.local_count = inner_len_slot + 1;
    chunk.emit_op_u16(Op::LOCAL_GET, inner_slot, line);
    host::emit(chunk, "wasm:js-string", "length", 1, line);
    chunk.emit_op_u16(Op::LOCAL_SET, inner_len_slot, line);

    // idx = parseInt(inner) — parseInt stops at first non-digit, so the
    // optional `,W` / `:fmt` suffixes are naturally trimmed.
    let idx_slot = chunk.local_count;
    chunk.local_count = idx_slot + 1;
    let pf_idx = chunk.add_import("ecma:number", "parseInt");
    chunk.emit_op_u16(Op::LOCAL_GET, inner_slot, line);
    chunk.emit_op_u16(Op::CALL_IMPORT, pf_idx, line);
    chunk.emit(1, line);
    chunk.emit_op(Op::I32_FROM_F64, line);
    chunk.emit_op_u16(Op::LOCAL_SET, idx_slot, line);

    let comma_slot = chunk.local_count;
    let colon_slot = comma_slot + 1;
    let format_slot = colon_slot + 1;
    let width_slot = format_slot + 1;
    chunk.local_count = width_slot + 1;

    chunk.emit_op_u16(Op::LOCAL_GET, inner_slot, line);
    chunk.emit_string_const(",", line);
    host::emit(chunk, "ecma:string", "indexOf", 2, line);
    chunk.emit_op_u16(Op::LOCAL_SET, comma_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, inner_slot, line);
    chunk.emit_string_const(":", line);
    host::emit(chunk, "ecma:string", "indexOf", 2, line);
    chunk.emit_op_u16(Op::LOCAL_SET, colon_slot, line);

    chunk.emit_string_const("", line);
    chunk.emit_op_u16(Op::LOCAL_SET, format_slot, line);

    chunk.emit_f64_const(0.0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, width_slot, line);

    let no_format_spec = chunk.emit_block(line);
    chunk.emit_op_u16(Op::LOCAL_GET, colon_slot, line);
    core_wasm::i32_const(chunk, line, 0);
    crate::emitter::ops::emit_dyn_lt(chunk, line);
    chunk.emit_br_if(0, line);
    chunk.emit_op_u16(Op::LOCAL_GET, inner_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, colon_slot, line);
    core_wasm::i32_const(chunk, line, 1);
    chunk.emit_op(Op::I32_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_GET, inner_len_slot, line);
    host::emit(chunk, "wasm:js-string", "substring", 3, line);
    chunk.emit_op_u16(Op::LOCAL_SET, format_slot, line);
    chunk.emit_end(line);
    chunk.patch_block(no_format_spec);

    let no_width_spec = chunk.emit_block(line);
    chunk.emit_op_u16(Op::LOCAL_GET, comma_slot, line);
    core_wasm::i32_const(chunk, line, 0);
    crate::emitter::ops::emit_dyn_lt(chunk, line);
    chunk.emit_br_if(0, line);
    let width_end_slot = chunk.local_count;
    chunk.local_count = width_end_slot + 1;
    chunk.emit_op_u16(Op::LOCAL_GET, inner_len_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, width_end_slot, line);
    let no_colon_after_width = chunk.emit_block(line);
    chunk.emit_op_u16(Op::LOCAL_GET, colon_slot, line);
    core_wasm::i32_const(chunk, line, 0);
    crate::emitter::ops::emit_dyn_lt(chunk, line);
    chunk.emit_br_if(0, line);
    chunk.emit_op_u16(Op::LOCAL_GET, colon_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, width_end_slot, line);
    chunk.emit_end(line);
    chunk.patch_block(no_colon_after_width);
    chunk.emit_op_u16(Op::LOCAL_GET, inner_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, comma_slot, line);
    core_wasm::i32_const(chunk, line, 1);
    chunk.emit_op(Op::I32_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_GET, width_end_slot, line);
    host::emit(chunk, "wasm:js-string", "substring", 3, line);
    chunk.emit_op_u16(Op::CALL_IMPORT, pf_idx, line);
    chunk.emit(1, line);
    chunk.emit_op_u16(Op::LOCAL_SET, width_slot, line);
    chunk.emit_end(line);
    chunk.patch_block(no_width_spec);

    // out = out + format(args[idx], format, width)
    chunk.emit_op_u16(Op::LOCAL_GET, out_slot, line);
    emit_dotnet_format_value_call(chunk, args_slot, idx_slot, format_slot, width_slot, line);
    crate::emitter::ops::emit_dyn_add(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_SET, out_slot, line);

    // i = end + 1 (skip past the closing '}')
    chunk.emit_op_u16(Op::LOCAL_GET, end_slot, line);
    core_wasm::i32_const(chunk, line, 1);
    chunk.emit_op(Op::I32_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_SET, i_slot, line);
}

/// Handle `}` at fmt[i]: either the `}}` escape or a stray `}` (which we
/// ignore by simply advancing past it — same as JS `dotnet_format` polyfill).
fn emit_handle_close_brace(
    chunk: &mut Chunk,
    fmt_slot: u16,
    i_slot: u16,
    len_slot: u16,
    out_slot: u16,
    line: u32,
) {
    let escape_block = chunk.emit_block(line);
    chunk.emit_op_u16(Op::LOCAL_GET, i_slot, line);
    core_wasm::i32_const(chunk, line, 1);
    chunk.emit_op(Op::I32_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_GET, len_slot, line);
    crate::emitter::ops::emit_dyn_lt(chunk, line);
    crate::emitter::ops::emit_dyn_not(chunk, line);
    chunk.emit_br_if(0, line);

    chunk.emit_op_u16(Op::LOCAL_GET, fmt_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, i_slot, line);
    core_wasm::i32_const(chunk, line, 1);
    chunk.emit_op(Op::I32_ADD, line);
    host::emit(chunk, "wasm:js-string", "charCodeAt", 2, line);
    chunk.emit_i32_const(b'}' as i32, line);
    crate::emitter::ops::emit_dyn_eq(chunk, line);
    crate::emitter::ops::emit_dyn_not(chunk, line);
    chunk.emit_br_if(0, line);

    // It IS `}}`: append '}' to out, advance i by 2.
    chunk.emit_op_u16(Op::LOCAL_GET, out_slot, line);
    push_const(chunk, Value::String(Arc::from("}")), line);
    crate::emitter::ops::emit_dyn_add(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_SET, out_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunk.emit_i32_const(2, line);
    chunk.emit_op(Op::I32_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_SET, i_slot, line);
    // br depth 2 = continue outer loop. Depths inside escape_block:
    // 0=escape_block, 1=close_block, 2=loop, 3=outer_block.
    chunk.emit_br(2, line);
    chunk.emit_end(line);
    chunk.patch_block(escape_block);

    // Stray `}` — just skip it.
    chunk.emit_op_u16(Op::LOCAL_GET, i_slot, line);
    core_wasm::i32_const(chunk, line, 1);
    chunk.emit_op(Op::I32_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_SET, i_slot, line);
}
