//! Inline-bytecode sprintf emitter.
//!
//! Called via `emit = "common:sprintf.format"` in language profiles.
//! Routes through `emitter/dispatch.rs` — NOT via stdlib globals.
//!
//! `emit_sprintf` adds the sprintf helper chunk to the program's chunk
//! list the first time it is called (identified by chunk name), then
//! emits a direct call-by-name to it. This keeps the implementation in
//! Rust bytecode (no JS polyfill) and in the proper emitter path.

use std::sync::Arc;
use vybe_bytecode::opcode::Op;
use vybe_bytecode::{Chunk, Value};

const CHUNK_NAME: &str = "__fmt_sprintf";

/// Called by `emitter/dispatch.rs` when a profile entry uses
/// `emit = "common:sprintf.format"`.
///
/// Stack convention (caller already pushed):
///   arg0 = fmt  (string)
///   arg1..argN = variadic args (N = argc - 1)
///
/// Because sprintf is variadic, we first collect args 1..N into an array,
/// then call the helper chunk that takes (fmt, args_array).
pub fn emit_sprintf(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    // Ensure the helper chunk is in the chunk list.
    let helper_idx = ensure_chunk(chunks);

    // The helper expects (fmt, args_array).  The compiler pushed the args
    // on the stack in order: [fmt, arg0, arg1, ...].
    // We need to pack args 1..N into an array first.
    let push_idx = chunks[0].add_import("ecma:array", "push");

    // Allocate a local to hold the args array being built.
    let arr_slot = chunks[current].local_count;
    chunks[current].local_count += 1;
    // Allocate a local to hold the format string.
    let fmt_slot = chunks[current].local_count;
    chunks[current].local_count += 1;

    let nargs = argc as i32; // total args including fmt
    let nrest = nargs - 1; // args after fmt

    // Pop args in reverse, store in temporaries, then rebuild.
    // Simpler: all args are already on the stack in forward order.
    // Strategy:
    //   1. Store all args into locals (right-to-left to keep stack sane).
    //   2. Build the args array from the stored locals.
    //   3. GLOBAL_GET the sprintf fn, call(fmt, args_array).

    // Allocate temp locals for the variadic args.
    let first_arg_slot = chunks[current].local_count;
    chunks[current].local_count += nrest.max(0) as u16;

    // Store variadic args (they are on top of stack, in order arg0..argN-1).
    // Stack order: ... fmt arg0 arg1 ... argN-1  (argN-1 on top)
    // Store from top to bottom.
    for k in (0..nrest).rev() {
        let slot = first_arg_slot + k as u16;
        chunks[current].emit_op_u16(Op::LOCAL_SET, slot, line);
        chunks[current].emit_op(Op::DROP, line);
    }
    // Now fmt is on top — store it.
    chunks[current].emit_op_u16(Op::LOCAL_SET, fmt_slot, line);
    chunks[current].emit_op(Op::DROP, line);

    // Build args array: []
    crate::emitter::collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, arr_slot, line);
    chunks[current].emit_op(Op::DROP, line);

    // Push each variadic arg into the array.
    for k in 0..nrest {
        let slot = first_arg_slot + k as u16;
        chunks[current].emit_op_u16(Op::LOCAL_GET, arr_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
        chunks[current].emit_op_u16(Op::CALL_IMPORT, push_idx, line);
        chunks[current].emit(2u8, line);
        chunks[current].emit_op(Op::DROP, line);
    }

    // Call the helper directly by function reference.
    chunks[current].emit_op_u16(Op::REF_FUNC, helper_idx as u16, line);
    chunks[current].emit(0u8, line); // upvalue count
    chunks[current].emit_op_u16(Op::LOCAL_GET, fmt_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    chunks[current].emit_op(Op::CALL_REF, line);
    chunks[current].emit(2u8, line);
}

/// Emit a direct sprintf helper call when the caller already has
/// `(fmt, args_array)` on the stack.
pub fn emit_sprintf_from_array(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let helper_idx = ensure_chunk(chunks);

    // CALL_REF expects [callee, arg0, arg1]. The caller currently has
    // [fmt, args_array], so stash and rebuild in the expected order.
    let arr_slot = chunks[current].local_count;
    chunks[current].local_count += 1;
    let fmt_slot = chunks[current].local_count;
    chunks[current].local_count += 1;

    chunks[current].emit_op_u16(Op::LOCAL_SET, arr_slot, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, fmt_slot, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_op_u16(Op::REF_FUNC, helper_idx as u16, line);
    chunks[current].emit(0u8, line); // upvalue count
    chunks[current].emit_op_u16(Op::LOCAL_GET, fmt_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    chunks[current].emit_op(Op::CALL_REF, line);
    chunks[current].emit(2u8, line);
}

/// Ensure the helper chunk exists in the chunk list.  Uses the chunk name
/// as the unique key — idempotent if called multiple times per compilation.
fn ensure_chunk(chunks: &mut Vec<Chunk>) -> usize {
    if let Some(idx) = chunks.iter().position(|c| c.name == CHUNK_NAME) {
        return idx;
    }
    // Build the imports chunk (chunk 0 is the module-level import table).
    let imports = &mut chunks[0];
    let helper = build_sprintf(imports);
    chunks.push(helper);
    chunks.len() - 1
}

const SSCANF_CHUNK: &str = "__fmt_sscanf";

/// `sscanf(string, format, ...)` — array-return form. The caller pushed
/// `[string, format]` plus any extra reference args. Returns an array of the
/// parsed values. (The by-reference assignment form is not supported; extra
/// args are dropped.)
pub fn emit_sscanf(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    if argc < 2 {
        for _ in 0..argc {
            chunks[current].emit_op(Op::DROP, line);
        }
        let k = chunks[current].add_constant(Value::Null);
        chunks[current].emit_op_u16(Op::CONST, k, line);
        return;
    }
    // Drop any extra reference args, leaving [string, format].
    for _ in 2..argc {
        chunks[current].emit_op(Op::DROP, line);
    }

    let helper_idx = ensure_sscanf_chunk(chunks);

    // CALL_REF expects [callee, input, fmt]; caller has [input, fmt].
    let fmt_slot = chunks[current].local_count;
    chunks[current].local_count += 1;
    let inp_slot = chunks[current].local_count;
    chunks[current].local_count += 1;

    chunks[current].emit_op_u16(Op::LOCAL_SET, fmt_slot, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, inp_slot, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_op_u16(Op::REF_FUNC, helper_idx as u16, line);
    chunks[current].emit(0u8, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, inp_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, fmt_slot, line);
    chunks[current].emit_op(Op::CALL_REF, line);
    chunks[current].emit(2u8, line);
}

fn ensure_sscanf_chunk(chunks: &mut Vec<Chunk>) -> usize {
    if let Some(idx) = chunks.iter().position(|c| c.name == SSCANF_CHUNK) {
        return idx;
    }
    let imports = &mut chunks[0];
    let helper = build_sscanf(imports);
    chunks.push(helper);
    chunks.len() - 1
}

// ── local slot indices inside the helper chunk ────────────────────────────
const FMT: u16 = 0;
const ARGS: u16 = 1;
const I: u16 = 2;
const FLEN: u16 = 3;
const OUT: u16 = 4;
const AIDX: u16 = 5;
const CH: u16 = 6;
const FLEFT: u16 = 7;
const FSIGN: u16 = 8;
const FZERO: u16 = 9;
const FSPACE: u16 = 10;
const FALT: u16 = 11;
const WIDTH: u16 = 12;
const PREC: u16 = 13;
const ARG: u16 = 14;
const RAW: u16 = 15;
const N: u16 = 16;
const CONV: u16 = 17;
const PADCH: u16 = 18;
const CUSTPAD: u16 = 19; // custom pad char from `%'X`, or "" if none
const FPAD: u16 = 20; // 1 when a custom pad char was given
const POS: u16 = 21; // positional arg number from `%N$` (scratch)
const SAVEI: u16 = 22; // saved `I` for positional-arg rewind (scratch)
const NLOCALS: u16 = 23;

fn lg(c: &mut Chunk, s: u16) {
    c.emit_op_u16(Op::LOCAL_GET, s, 0);
}
fn ls(c: &mut Chunk, s: u16) {
    c.emit_op_u16(Op::LOCAL_SET, s, 0);
    c.emit_op(Op::DROP, 0);
}
fn ci(c: &mut Chunk, v: i32) {
    let k = c.add_constant(Value::I32(v));
    c.emit_op_u16(Op::CONST, k, 0);
}
fn cf(c: &mut Chunk, v: f64) {
    let k = c.add_constant(Value::F64(v));
    c.emit_op_u16(Op::CONST, k, 0);
}
fn cs(c: &mut Chunk, v: &str) {
    let k = c.add_constant(Value::String(Arc::from(v)));
    c.emit_op_u16(Op::CONST, k, 0);
}
fn hc(c: &mut Chunk, i: u16, a: u8) {
    c.emit_op_u16(Op::CALL_IMPORT, i, 0);
    c.emit(a, 0);
}
fn inc(c: &mut Chunk, s: u16) {
    lg(c, s);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::I32_ADD, 0);
    ls(c, s);
}

/// Build the sprintf helper chunk.  Takes (fmt: string, args: array) and
/// returns the formatted string.
pub fn build_sprintf(_imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new(CHUNK_NAME);
    c.arity = 2;
    c.local_count = NLOCALS;

    let str_len = c.add_import("ecma:string", "length");
    let str_ccat = c.add_import("ecma:string", "charCodeAt");
    let str_chat = c.add_import("ecma:string", "charAt");
    let str_slice = c.add_import("ecma:string", "slice");
    let str_tostr = c.add_import("ecma:string", "String");
    let str_fcc = c.add_import("ecma:string", "fromCharCode");
    let str_cat = c.add_import("ecma:string", "concat");
    let str_upper = c.add_import("ecma:string", "toUpperCase");
    let str_pstart = c.add_import("ecma:string", "padStart");
    let str_pend = c.add_import("ecma:string", "padEnd");
    let num_num = c.add_import("ecma:number", "Number");
    let num_fixed = c.add_import("ecma:number", "toFixed");
    let num_exp = c.add_import("ecma:number", "toExponential");
    let num_radix = c.add_import("ecma:number", "toString");
    let num_prec = c.add_import("ecma:number", "toPrecision");
    let math_trunc = c.add_import("ecma:math", "trunc");
    let math_abs = c.add_import("ecma:math", "abs");
    let math_pow = c.add_import("ecma:math", "pow");
    let arr_at = c.add_import("ecma:array", "at");

    // init
    lg(&mut c, FMT);
    hc(&mut c, str_len, 1);
    ls(&mut c, FLEN);
    cs(&mut c, "");
    ls(&mut c, OUT);
    c.emit_op(Op::I32_CONST_0, 0);
    ls(&mut c, I);
    c.emit_op(Op::I32_CONST_0, 0);
    ls(&mut c, AIDX);

    // outer: block $ob, loop $ol
    let ob = c.emit_block(0);
    let (ol, _) = c.emit_loop_s(0);
    lg(&mut c, I);
    lg(&mut c, FLEN);
    c.emit_op(Op::I32_GE_S, 0);
    c.emit_br_if(1, 0);

    // read ch
    lg(&mut c, FMT);
    lg(&mut c, I);
    hc(&mut c, str_ccat, 2);
    c.emit_op(Op::I32_FROM_F64, 0);
    ls(&mut c, CH);

    // non-% path (depth: plain_blk=0, ol=1, ob=2)
    {
        let pb = c.emit_block(0);
        lg(&mut c, CH);
        ci(&mut c, 37);
        c.emit_op(Op::I32_EQ, 0);
        c.emit_br_if(0, 0);
        lg(&mut c, OUT);
        lg(&mut c, FMT);
        lg(&mut c, I);
        hc(&mut c, str_chat, 2);
        hc(&mut c, str_cat, 2);
        ls(&mut c, OUT);
        inc(&mut c, I);
        c.emit_br(1, 0); // continue ol
        c.emit_end(0);
        c.patch_block(pb);
    }
    inc(&mut c, I); // skip '%'

    // Fast paths for the two most common simple cases. This avoids
    // entering the full flag/width/precision parser for plain "%%" and
    // "%c" specifiers.
    {
        let fp = c.emit_block(0);
        lg(&mut c, I);
        lg(&mut c, FLEN);
        c.emit_op(Op::I32_GE_S, 0);
        c.emit_br_if(0, 0);

        lg(&mut c, FMT);
        lg(&mut c, I);
        hc(&mut c, str_ccat, 2);
        c.emit_op(Op::I32_FROM_F64, 0);
        ls(&mut c, CH);

        // %% -> literal percent
        {
            let pct = c.emit_block(0);
            lg(&mut c, CH);
            ci(&mut c, 37);
            c.emit_op(Op::I32_NE, 0);
            c.emit_br_if(0, 0);
            lg(&mut c, OUT);
            cs(&mut c, "%");
            hc(&mut c, str_cat, 2);
            ls(&mut c, OUT);
            inc(&mut c, I);
            c.emit_br(2, 0); // continue outer loop
            c.emit_end(0);
            c.patch_block(pct);
        }

        // %c -> next arg as character
        {
            let chr = c.emit_block(0);
            lg(&mut c, CH);
            ci(&mut c, 99);
            c.emit_op(Op::I32_NE, 0);
            c.emit_br_if(0, 0);
            lg(&mut c, ARGS);
            lg(&mut c, AIDX);
            hc(&mut c, arr_at, 2);
            ls(&mut c, ARG);
            inc(&mut c, AIDX);

            // If arg is already a string, take its first character.
            // Otherwise coerce to char code and convert via fromCharCode.
            lg(&mut c, ARG);
            c.emit_op(Op::REF_IS_STRING, 0);
            let line = 0;
            c.emit_if(line);
            lg(&mut c, ARG);
            ci(&mut c, 0);
            hc(&mut c, str_chat, 2);
            ls(&mut c, RAW);
            c.emit_else(line);
            lg(&mut c, ARG);
            hc(&mut c, num_num, 1);
            hc(&mut c, str_fcc, 1);
            ls(&mut c, RAW);
            c.emit_end(line);

            lg(&mut c, OUT);
            lg(&mut c, RAW);
            hc(&mut c, str_cat, 2);
            ls(&mut c, OUT);
            inc(&mut c, I);
            c.emit_br(2, 0); // continue outer loop
            c.emit_end(0);
            c.patch_block(chr);
        }

        c.emit_end(0);
        c.patch_block(fp);
    }

    // positional argument: `%N$...` selects argument N (1-based). Parse the
    // leading digit run; if it is followed by '$' it is an explicit arg
    // index (set AIDX = N-1), otherwise rewind so the digits parse as width.
    {
        let pos_block = c.emit_block(0);
        lg(&mut c, I);
        ls(&mut c, SAVEI);
        c.emit_op(Op::I32_CONST_0, 0);
        ls(&mut c, POS);
        // scan digits into POS
        {
            let pl = c.emit_block(0);
            let (plp, _) = c.emit_loop_s(0);
            lg(&mut c, I);
            lg(&mut c, FLEN);
            c.emit_op(Op::I32_GE_S, 0);
            c.emit_br_if(1, 0);
            lg(&mut c, FMT);
            lg(&mut c, I);
            hc(&mut c, str_ccat, 2);
            c.emit_op(Op::I32_FROM_F64, 0);
            ls(&mut c, CH);
            digit_or_break(&mut c, CH, 2); // exit pl when CH is not a digit
            lg(&mut c, POS);
            ci(&mut c, 10);
            c.emit_op(Op::I32_MUL, 0);
            lg(&mut c, CH);
            ci(&mut c, 48);
            c.emit_op(Op::I32_SUB, 0);
            c.emit_op(Op::I32_ADD, 0);
            ls(&mut c, POS);
            inc(&mut c, I);
            c.emit_br(0, 0);
            c.emit_end(0);
            c.patch_loop(plp);
            c.emit_end(0);
            c.patch_block(pl);
        }
        // if next char is '$' → use POS as arg index; else rewind I
        {
            let nd = c.emit_block(0);
            lg(&mut c, I);
            lg(&mut c, FLEN);
            c.emit_op(Op::I32_GE_S, 0);
            c.emit_br_if(0, 0);
            lg(&mut c, FMT);
            lg(&mut c, I);
            hc(&mut c, str_ccat, 2);
            c.emit_op(Op::I32_FROM_F64, 0);
            ls(&mut c, CH);
            lg(&mut c, CH);
            ci(&mut c, 36);
            c.emit_op(Op::I32_NE, 0);
            c.emit_br_if(0, 0);
            // dollar present: AIDX = POS - 1 ; consume '$'
            lg(&mut c, POS);
            ci(&mut c, 1);
            c.emit_op(Op::I32_SUB, 0);
            ls(&mut c, AIDX);
            inc(&mut c, I);
            c.emit_br(1, 0); // exit pos_block, skipping the rewind
            c.emit_end(0);
            c.patch_block(nd);
        }
        lg(&mut c, SAVEI);
        ls(&mut c, I); // rewind: digits were width, not arg index
        c.emit_end(0);
        c.patch_block(pos_block);
    }

    // reset flags
    c.emit_op(Op::I32_CONST_0, 0);
    ls(&mut c, FLEFT);
    c.emit_op(Op::I32_CONST_0, 0);
    ls(&mut c, FSIGN);
    c.emit_op(Op::I32_CONST_0, 0);
    ls(&mut c, FZERO);
    c.emit_op(Op::I32_CONST_0, 0);
    ls(&mut c, FSPACE);
    c.emit_op(Op::I32_CONST_0, 0);
    ls(&mut c, FALT);
    c.emit_op(Op::I32_CONST_0, 0);
    ls(&mut c, WIDTH);
    c.emit_op(Op::I32_CONST_0, 0);
    ls(&mut c, FPAD);
    cs(&mut c, "");
    ls(&mut c, CUSTPAD);
    ci(&mut c, -1);
    ls(&mut c, PREC);

    // spec block — "continue outer" from inside spec = br to ol
    // depths inside spec: spec=0, ol=1, ob=2
    let spec = c.emit_block(0);

    // flag loop
    {
        let fl = c.emit_block(0);
        let (flp, _) = c.emit_loop_s(0);
        lg(&mut c, I);
        lg(&mut c, FLEN);
        c.emit_op(Op::I32_GE_S, 0);
        c.emit_br_if(1, 0);
        lg(&mut c, FMT);
        lg(&mut c, I);
        hc(&mut c, str_ccat, 2);
        c.emit_op(Op::I32_FROM_F64, 0);
        ls(&mut c, CH);
        flag_arm(&mut c, CH, 45, FLEFT);
        flag_arm(&mut c, CH, 43, FSIGN);
        flag_arm(&mut c, CH, 32, FSPACE);
        flag_arm(&mut c, CH, 35, FALT);
        flag_arm(&mut c, CH, 48, FZERO);
        // custom pad char: `'X` takes the next char as the pad character.
        {
            let b = c.emit_block(0);
            lg(&mut c, CH);
            ci(&mut c, 39);
            c.emit_op(Op::I32_NE, 0);
            c.emit_br_if(0, 0);
            inc(&mut c, I); // skip the quote
            lg(&mut c, FMT);
            lg(&mut c, I);
            hc(&mut c, str_chat, 2);
            ls(&mut c, CUSTPAD);
            ci(&mut c, 1);
            ls(&mut c, FPAD);
            inc(&mut c, I); // skip the pad char
            c.emit_br(1, 0); // continue flag loop
            c.emit_end(0);
            c.patch_block(b);
        }
        c.emit_br(1, 0); // exit fl
        c.emit_end(0);
        c.patch_loop(flp);
        c.emit_end(0);
        c.patch_block(fl);
    }
    read_ch(&mut c, FMT, I, FLEN, CH, str_ccat);

    // width loop
    {
        let wl = c.emit_block(0);
        let (wlp, _) = c.emit_loop_s(0);
        lg(&mut c, I);
        lg(&mut c, FLEN);
        c.emit_op(Op::I32_GE_S, 0);
        c.emit_br_if(1, 0);
        lg(&mut c, FMT);
        lg(&mut c, I);
        hc(&mut c, str_ccat, 2);
        c.emit_op(Op::I32_FROM_F64, 0);
        ls(&mut c, CH);
        digit_or_break(&mut c, CH, 2); // 2 = exit wl (from inside digit_or_break's block)
        lg(&mut c, WIDTH);
        ci(&mut c, 10);
        c.emit_op(Op::I32_MUL, 0);
        lg(&mut c, CH);
        ci(&mut c, 48);
        c.emit_op(Op::I32_SUB, 0);
        c.emit_op(Op::I32_ADD, 0);
        ls(&mut c, WIDTH);
        inc(&mut c, I);
        c.emit_br(0, 0);
        c.emit_end(0);
        c.patch_loop(wlp);
        c.emit_end(0);
        c.patch_block(wl);
    }
    read_ch(&mut c, FMT, I, FLEN, CH, str_ccat);

    // precision
    {
        let dot = c.emit_block(0);
        lg(&mut c, I);
        lg(&mut c, FLEN);
        c.emit_op(Op::I32_GE_S, 0);
        c.emit_br_if(0, 0);
        lg(&mut c, CH);
        ci(&mut c, 46);
        c.emit_op(Op::I32_NE, 0);
        c.emit_br_if(0, 0);
        inc(&mut c, I);
        c.emit_op(Op::I32_CONST_0, 0);
        ls(&mut c, PREC);
        let pl = c.emit_block(0);
        let (plp, _) = c.emit_loop_s(0);
        lg(&mut c, I);
        lg(&mut c, FLEN);
        c.emit_op(Op::I32_GE_S, 0);
        c.emit_br_if(1, 0);
        lg(&mut c, FMT);
        lg(&mut c, I);
        hc(&mut c, str_ccat, 2);
        c.emit_op(Op::I32_FROM_F64, 0);
        ls(&mut c, CH);
        digit_or_break(&mut c, CH, 2);
        lg(&mut c, PREC);
        ci(&mut c, 10);
        c.emit_op(Op::I32_MUL, 0);
        lg(&mut c, CH);
        ci(&mut c, 48);
        c.emit_op(Op::I32_SUB, 0);
        c.emit_op(Op::I32_ADD, 0);
        ls(&mut c, PREC);
        inc(&mut c, I);
        c.emit_br(0, 0);
        c.emit_end(0);
        c.patch_loop(plp);
        c.emit_end(0);
        c.patch_block(pl);
        read_ch(&mut c, FMT, I, FLEN, CH, str_ccat);
        c.emit_end(0);
        c.patch_block(dot);
    }

    // C/PHP length modifiers (`h`, `hh`, `l`, `ll`, `j`, `z`, `t`, `L`) do
    // not change the dynamic value representation here; skip them so `%ld`
    // still dispatches as `%d`.
    {
        let lm = c.emit_block(0);
        let (lmp, _) = c.emit_loop_s(0);
        lg(&mut c, I);
        lg(&mut c, FLEN);
        c.emit_op(Op::I32_GE_S, 0);
        c.emit_br_if(1, 0);
        lg(&mut c, FMT);
        lg(&mut c, I);
        hc(&mut c, str_ccat, 2);
        c.emit_op(Op::I32_FROM_F64, 0);
        ls(&mut c, CH);

        lg(&mut c, CH);
        ci(&mut c, 104);
        c.emit_op(Op::I32_EQ, 0); // h
        lg(&mut c, CH);
        ci(&mut c, 108);
        c.emit_op(Op::I32_EQ, 0); // l
        c.emit_op(Op::I32_OR, 0);
        lg(&mut c, CH);
        ci(&mut c, 106);
        c.emit_op(Op::I32_EQ, 0); // j
        c.emit_op(Op::I32_OR, 0);
        lg(&mut c, CH);
        ci(&mut c, 122);
        c.emit_op(Op::I32_EQ, 0); // z
        c.emit_op(Op::I32_OR, 0);
        lg(&mut c, CH);
        ci(&mut c, 116);
        c.emit_op(Op::I32_EQ, 0); // t
        c.emit_op(Op::I32_OR, 0);
        lg(&mut c, CH);
        ci(&mut c, 76);
        c.emit_op(Op::I32_EQ, 0); // L
        c.emit_op(Op::I32_OR, 0);
        c.emit_op(Op::I32_EQZ, 0);
        c.emit_br_if(1, 0);

        inc(&mut c, I);
        c.emit_br(0, 0);
        c.emit_end(0);
        c.patch_loop(lmp);
        c.emit_end(0);
        c.patch_block(lm);
    }

    // read conversion char; if past end emit "%" and break loop
    {
        let eof = c.emit_block(0);
        lg(&mut c, I);
        lg(&mut c, FLEN);
        c.emit_op(Op::I32_LT_S, 0);
        c.emit_br_if(0, 0);
        lg(&mut c, OUT);
        cs(&mut c, "%");
        hc(&mut c, str_cat, 2);
        ls(&mut c, OUT);
        c.emit_br(3, 0); // break ob
        c.emit_end(0);
        c.patch_block(eof);
    }
    lg(&mut c, FMT);
    lg(&mut c, I);
    hc(&mut c, str_ccat, 2);
    c.emit_op(Op::I32_FROM_F64, 0);
    ls(&mut c, CONV);
    inc(&mut c, I);

    // %% — depths: pct=0, spec=1, ol=2  → continue ol = br 2
    {
        let pct = c.emit_block(0);
        lg(&mut c, CONV);
        ci(&mut c, 37);
        c.emit_op(Op::I32_NE, 0);
        c.emit_br_if(0, 0);
        lg(&mut c, OUT);
        cs(&mut c, "%");
        hc(&mut c, str_cat, 2);
        ls(&mut c, OUT);
        c.emit_br(2, 0);
        c.emit_end(0);
        c.patch_block(pct);
    }

    // load arg
    lg(&mut c, ARGS);
    lg(&mut c, AIDX);
    hc(&mut c, arr_at, 2);
    ls(&mut c, ARG);
    inc(&mut c, AIDX);
    lg(&mut c, ARG);
    hc(&mut c, num_num, 1);
    ls(&mut c, N);
    cs(&mut c, "");
    ls(&mut c, RAW);

    // conversions
    conv_s(&mut c, str_tostr, str_slice);
    conv_d(&mut c, math_trunc, math_abs, num_radix, str_cat);
    conv_u(&mut c, num_radix);
    conv_f(&mut c, num_fixed, str_cat, math_pow);
    conv_e(&mut c, num_exp, str_upper, str_cat);
    conv_radix(
        &mut c, math_trunc, num_radix, str_upper, str_cat, 120, 16, false, "0x",
    );
    conv_radix(
        &mut c, math_trunc, num_radix, str_upper, str_cat, 88, 16, true, "0X",
    );
    conv_radix(
        &mut c, math_trunc, num_radix, str_upper, str_cat, 111, 8, false, "0",
    );
    conv_radix(
        &mut c, math_trunc, num_radix, str_upper, str_cat, 98, 2, false, "",
    );
    conv_g(&mut c, num_prec, str_upper, str_cat);
    conv_c(&mut c, num_num, str_fcc);

    // pad char: custom `'X` wins; else if fzero && !fleft → "0" else " "
    {
        lg(&mut c, FPAD);
        c.emit_if_value(0);
        lg(&mut c, CUSTPAD);
        c.emit_else(0);
        lg(&mut c, FZERO);
        lg(&mut c, FLEFT);
        c.emit_op(Op::I32_EQZ, 0);
        c.emit_op(Op::I32_AND, 0);
        c.emit_if_value(0);
        cs(&mut c, "0");
        c.emit_else(0);
        cs(&mut c, " ");
        c.emit_end(0);
        c.emit_end(0);
    }
    ls(&mut c, PADCH);

    // apply width
    {
        let pad = c.emit_block(0);
        lg(&mut c, WIDTH);
        c.emit_op(Op::I32_CONST_0, 0);
        c.emit_op(Op::I32_LE_S, 0);
        c.emit_br_if(0, 0);
        {
            let la = c.emit_block(0);
            lg(&mut c, FLEFT);
            c.emit_op(Op::I32_EQZ, 0);
            c.emit_br_if(0, 0);
            lg(&mut c, RAW);
            lg(&mut c, WIDTH);
            c.emit_op(Op::F64_FROM_I32, 0);
            lg(&mut c, PADCH);
            hc(&mut c, str_pend, 3);
            ls(&mut c, RAW);
            c.emit_br(1, 0);
            c.emit_end(0);
            c.patch_block(la);
        }
        {
            let za = c.emit_block(0);
            lg(&mut c, FLEFT);
            c.emit_br_if(0, 0);
            lg(&mut c, PADCH);
            ci(&mut c, 0);
            c.emit_op(Op::STR_CHAR_CODE_AT, 0);
            ci(&mut c, 48);
            c.emit_op(Op::I32_NE, 0);
            c.emit_br_if(0, 0);
            lg(&mut c, RAW);
            c.emit_op(Op::STR_LENGTH, 0);
            c.emit_op(Op::I32_EQZ, 0);
            c.emit_br_if(0, 0);
            lg(&mut c, RAW);
            ci(&mut c, 0);
            c.emit_op(Op::STR_CHAR_CODE_AT, 0);
            ci(&mut c, 43);
            c.emit_op(Op::I32_EQ, 0);
            lg(&mut c, RAW);
            ci(&mut c, 0);
            c.emit_op(Op::STR_CHAR_CODE_AT, 0);
            ci(&mut c, 45);
            c.emit_op(Op::I32_EQ, 0);
            c.emit_op(Op::I32_OR, 0);
            lg(&mut c, RAW);
            ci(&mut c, 0);
            c.emit_op(Op::STR_CHAR_CODE_AT, 0);
            ci(&mut c, 32);
            c.emit_op(Op::I32_EQ, 0);
            c.emit_op(Op::I32_OR, 0);
            c.emit_op(Op::I32_EQZ, 0);
            c.emit_br_if(0, 0);
            lg(&mut c, RAW);
            ci(&mut c, 0);
            c.emit_op(Op::STR_CHAR_AT, 0);
            ls(&mut c, CUSTPAD);
            lg(&mut c, CUSTPAD);
            lg(&mut c, RAW);
            ci(&mut c, 1);
            lg(&mut c, RAW);
            c.emit_op(Op::STR_LENGTH, 0);
            c.emit_op(Op::STR_SLICE, 0);
            lg(&mut c, WIDTH);
            ci(&mut c, 1);
            c.emit_op(Op::I32_SUB, 0);
            c.emit_op(Op::F64_FROM_I32, 0);
            cs(&mut c, "0");
            hc(&mut c, str_pstart, 3);
            hc(&mut c, str_cat, 2);
            ls(&mut c, RAW);
            c.emit_br(1, 0);
            c.emit_end(0);
            c.patch_block(za);
        }
        lg(&mut c, RAW);
        lg(&mut c, WIDTH);
        c.emit_op(Op::F64_FROM_I32, 0);
        lg(&mut c, PADCH);
        hc(&mut c, str_pstart, 3);
        ls(&mut c, RAW);
        c.emit_end(0);
        c.patch_block(pad);
    }

    lg(&mut c, OUT);
    lg(&mut c, RAW);
    hc(&mut c, str_cat, 2);
    ls(&mut c, OUT);

    c.emit_end(0);
    c.patch_block(spec);
    c.emit_br(0, 0); // continue ol
    c.emit_end(0);
    c.patch_loop(ol);
    c.emit_end(0);
    c.patch_block(ob);

    lg(&mut c, OUT);
    c.emit_op(Op::RETURN, 0);
    c
}

fn read_ch(c: &mut Chunk, fmt: u16, i: u16, flen: u16, ch: u16, ccat: u16) {
    let b = c.emit_block(0);
    lg(c, i);
    lg(c, flen);
    c.emit_op(Op::I32_GE_S, 0);
    c.emit_br_if(0, 0);
    lg(c, fmt);
    lg(c, i);
    hc(c, ccat, 2);
    c.emit_op(Op::I32_FROM_F64, 0);
    ls(c, ch);
    c.emit_end(0);
    c.patch_block(b);
}

// If CH not in '0'..'9', branch to break the caller's loop.
// `depth` is counted from inside this function's own block.
fn digit_or_break(c: &mut Chunk, ch: u16, depth: u32) {
    let b = c.emit_block(0);
    lg(c, ch);
    ci(c, 48);
    c.emit_op(Op::I32_GE_S, 0);
    lg(c, ch);
    ci(c, 57);
    c.emit_op(Op::I32_LE_S, 0);
    c.emit_op(Op::I32_AND, 0);
    c.emit_br_if(0, 0);
    c.emit_br(depth, 0);
    c.emit_end(0);
    c.patch_block(b);
}

fn flag_arm(c: &mut Chunk, ch: u16, code: i32, flag: u16) {
    let b = c.emit_block(0);
    lg(c, ch);
    ci(c, code);
    c.emit_op(Op::I32_NE, 0);
    c.emit_br_if(0, 0);
    ci(c, 1);
    ls(c, flag);
    inc(c, I);
    c.emit_br(1, 0); // continue flag loop (arm_blk=0, flp=1)
    c.emit_end(0);
    c.patch_block(b);
}

fn conv_s(c: &mut Chunk, str_tostr: u16, str_slice: u16) {
    let b = c.emit_block(0);
    lg(c, CONV);
    ci(c, 115);
    c.emit_op(Op::I32_NE, 0);
    c.emit_br_if(0, 0);
    lg(c, ARG);
    hc(c, str_tostr, 1);
    ls(c, RAW);
    {
        let pb = c.emit_block(0);
        lg(c, PREC);
        ci(c, 0);
        c.emit_op(Op::I32_LT_S, 0);
        c.emit_br_if(0, 0);
        lg(c, RAW);
        ci(c, 0);
        c.emit_op(Op::F64_FROM_I32, 0);
        lg(c, PREC);
        c.emit_op(Op::F64_FROM_I32, 0);
        hc(c, str_slice, 3);
        ls(c, RAW);
        c.emit_end(0);
        c.patch_block(pb);
    }
    c.emit_end(0);
    c.patch_block(b);
}

fn conv_d(c: &mut Chunk, math_trunc: u16, math_abs: u16, num_radix: u16, str_cat: u16) {
    let b = c.emit_block(0);
    lg(c, CONV);
    ci(c, 100);
    c.emit_op(Op::I32_EQ, 0);
    lg(c, CONV);
    ci(c, 105);
    c.emit_op(Op::I32_EQ, 0);
    c.emit_op(Op::I32_OR, 0);
    c.emit_op(Op::I32_EQZ, 0);
    c.emit_br_if(0, 0);
    lg(c, N);
    hc(c, math_trunc, 1);
    ls(c, N);
    lg(c, N);
    hc(c, math_abs, 1);
    cf(c, 10.0);
    hc(c, num_radix, 2);
    ls(c, RAW);
    suppress_zero_digits_for_zero_precision(c);
    // sign
    {
        let nb = c.emit_block(0);
        lg(c, N);
        cf(c, 0.0);
        c.emit_op(Op::F64_LT, 0);
        c.emit_br_if(0, 0);
        // non-negative: fsign?
        {
            let sb = c.emit_block(0);
            lg(c, FSIGN);
            c.emit_op(Op::I32_EQZ, 0);
            c.emit_br_if(0, 0);
            cs(c, "+");
            lg(c, RAW);
            hc(c, str_cat, 2);
            ls(c, RAW);
            // Exit the whole conv block (depth: sb=0, nb=1, b=2) so the
            // trailing "-" append after `end nb` is skipped. br(1) would
            // only exit nb and land on the "-", producing "-+42".
            c.emit_br(2, 0);
            c.emit_end(0);
            c.patch_block(sb);
        }
        // fspace?
        {
            let spb = c.emit_block(0);
            lg(c, FSPACE);
            c.emit_op(Op::I32_EQZ, 0);
            c.emit_br_if(0, 0);
            cs(c, " ");
            lg(c, RAW);
            hc(c, str_cat, 2);
            ls(c, RAW);
            c.emit_end(0);
            c.patch_block(spb);
        }
        c.emit_br(1, 0); // skip negative "-"
        c.emit_end(0);
        c.patch_block(nb);
    }
    cs(c, "-");
    lg(c, RAW);
    hc(c, str_cat, 2);
    ls(c, RAW);
    c.emit_end(0);
    c.patch_block(b);
}

fn conv_u(c: &mut Chunk, num_radix: u16) {
    let b = c.emit_block(0);
    lg(c, CONV);
    ci(c, 117);
    c.emit_op(Op::I32_NE, 0);
    c.emit_br_if(0, 0);
    // C `%u` expects a 32-bit unsigned view of the argument.
    // Keep positive values as-is; for negatives, add 2^32 (e.g. -1 -> 4294967295).
    lg(c, N);
    ls(c, SAVEI);
    lg(c, SAVEI);
    cf(c, 0.0);
    c.emit_op(Op::F64_LT, 0);
    c.emit_if_value(0);
    lg(c, SAVEI);
    cf(c, 4294967296.0);
    c.emit_op(Op::F64_ADD, 0);
    c.emit_else(0);
    lg(c, SAVEI);
    c.emit_end(0);
    cf(c, 10.0);
    hc(c, num_radix, 2);
    ls(c, RAW);
    suppress_zero_digits_for_zero_precision(c);
    c.emit_end(0);
    c.patch_block(b);
}

fn conv_f(c: &mut Chunk, num_fixed: u16, str_cat: u16, math_pow: u16) {
    let b = c.emit_block(0);
    lg(c, CONV);
    ci(c, 102);
    c.emit_op(Op::I32_EQ, 0);
    lg(c, CONV);
    ci(c, 70);
    c.emit_op(Op::I32_EQ, 0);
    c.emit_op(Op::I32_OR, 0);
    c.emit_op(Op::I32_EQZ, 0);
    c.emit_br_if(0, 0);

    cf(c, 10.0);
    push_precision_or_default(c, 6.0);
    hc(c, math_pow, 2);
    ls(c, SAVEI);

    lg(c, N);
    lg(c, SAVEI);
    c.emit_op(Op::F64_MUL, 0);
    c.emit_op(Op::F64_NEAREST, 0);
    lg(c, SAVEI);
    c.emit_op(Op::F64_DIV, 0);
    {
        lg(c, PREC);
        ci(c, 0);
        c.emit_op(Op::I32_GE_S, 0);
        c.emit_if_value(0);
        lg(c, PREC);
        c.emit_op(Op::F64_FROM_I32, 0);
        c.emit_else(0);
        cf(c, 6.0);
        c.emit_end(0);
    }
    hc(c, num_fixed, 2);
    ls(c, RAW);
    normalize_negative_zero_raw(c);
    sign_pos(c, str_cat);
    c.emit_end(0);
    c.patch_block(b);
}

fn push_precision_or_default(c: &mut Chunk, default_precision: f64) {
    lg(c, PREC);
    ci(c, 0);
    c.emit_op(Op::I32_GE_S, 0);
    c.emit_if_value(0);
    lg(c, PREC);
    c.emit_op(Op::F64_FROM_I32, 0);
    c.emit_else(0);
    cf(c, default_precision);
    c.emit_end(0);
}

fn normalize_negative_zero_raw(c: &mut Chunk) {
    let b = c.emit_block(0);
    lg(c, RAW);
    c.emit_op(Op::STR_LENGTH, 0);
    ci(c, 2);
    c.emit_op(Op::I32_LT_S, 0);
    c.emit_br_if(0, 0);
    lg(c, RAW);
    ci(c, 0);
    c.emit_op(Op::STR_CHAR_CODE_AT, 0);
    ci(c, 45);
    c.emit_op(Op::I32_NE, 0);
    c.emit_br_if(0, 0);
    lg(c, RAW);
    ci(c, 1);
    c.emit_op(Op::STR_CHAR_CODE_AT, 0);
    ci(c, 48);
    c.emit_op(Op::I32_NE, 0);
    c.emit_br_if(0, 0);
    lg(c, N);
    cf(c, -0.000000001);
    c.emit_op(Op::F64_GT, 0);
    c.emit_op(Op::I32_EQZ, 0);
    c.emit_br_if(0, 0);
    lg(c, RAW);
    ci(c, 1);
    lg(c, RAW);
    c.emit_op(Op::STR_LENGTH, 0);
    c.emit_op(Op::STR_SLICE, 0);
    ls(c, RAW);
    c.emit_end(0);
    c.patch_block(b);
}

fn suppress_zero_digits_for_zero_precision(c: &mut Chunk) {
    let b = c.emit_block(0);
    lg(c, PREC);
    c.emit_op(Op::I32_EQZ, 0);
    c.emit_op(Op::I32_EQZ, 0);
    c.emit_br_if(0, 0);
    lg(c, N);
    cf(c, 0.0);
    c.emit_op(Op::F64_EQ, 0);
    c.emit_op(Op::I32_EQZ, 0);
    c.emit_br_if(0, 0);
    cs(c, "");
    ls(c, RAW);
    c.emit_end(0);
    c.patch_block(b);
}

fn conv_e(c: &mut Chunk, num_exp: u16, str_upper: u16, str_cat: u16) {
    let b = c.emit_block(0);
    lg(c, CONV);
    ci(c, 101);
    c.emit_op(Op::I32_EQ, 0);
    lg(c, CONV);
    ci(c, 69);
    c.emit_op(Op::I32_EQ, 0);
    c.emit_op(Op::I32_OR, 0);
    c.emit_op(Op::I32_EQZ, 0);
    c.emit_br_if(0, 0);
    lg(c, N);
    {
        lg(c, PREC);
        ci(c, 0);
        c.emit_op(Op::I32_GE_S, 0);
        c.emit_if_value(0);
        lg(c, PREC);
        c.emit_op(Op::F64_FROM_I32, 0);
        c.emit_else(0);
        cf(c, 6.0);
        c.emit_end(0);
    }
    hc(c, num_exp, 2);
    ls(c, RAW);
    {
        let ub = c.emit_block(0);
        lg(c, CONV);
        ci(c, 69);
        c.emit_op(Op::I32_NE, 0);
        c.emit_br_if(0, 0);
        lg(c, RAW);
        hc(c, str_upper, 1);
        ls(c, RAW);
        c.emit_end(0);
        c.patch_block(ub);
    }
    sign_pos(c, str_cat);
    c.emit_end(0);
    c.patch_block(b);
}

// %g / %G — shortest of %e/%f at `precision` significant digits, with
// trailing zeros (and a bare trailing '.') removed. Uses Number.toPrecision
// then strips, which matches PHP for the common fixed-notation cases.
fn conv_g(c: &mut Chunk, num_prec: u16, str_upper: u16, str_cat: u16) {
    let b = c.emit_block(0);
    lg(c, CONV);
    ci(c, 103);
    c.emit_op(Op::I32_EQ, 0);
    lg(c, CONV);
    ci(c, 71);
    c.emit_op(Op::I32_EQ, 0);
    c.emit_op(Op::I32_OR, 0);
    c.emit_op(Op::I32_EQZ, 0);
    c.emit_br_if(0, 0);

    // RAW = N.toPrecision(prec) ; prec = PREC>=0 ? max(PREC,1) : 6
    lg(c, N);
    {
        lg(c, PREC);
        ci(c, 0);
        c.emit_op(Op::I32_LT_S, 0);
        c.emit_if_value(0);
        cf(c, 6.0);
        c.emit_else(0);
        lg(c, PREC);
        ci(c, 1);
        c.emit_op(Op::I32_LT_S, 0);
        c.emit_if_value(0);
        cf(c, 1.0);
        c.emit_else(0);
        lg(c, PREC);
        c.emit_op(Op::F64_FROM_I32, 0);
        c.emit_end(0);
        c.emit_end(0);
    }
    hc(c, num_prec, 2);
    ls(c, RAW);

    // Strip trailing zeros only for fixed notation (RAW has '.' and no 'e').
    {
        let sb = c.emit_block(0);
        lg(c, RAW);
        cs(c, ".");
        c.emit_op(Op::STR_INDEX_OF, 0);
        ci(c, 0);
        c.emit_op(Op::I32_LT_S, 0);
        c.emit_br_if(0, 0); // no '.' → skip
        lg(c, RAW);
        cs(c, "e");
        c.emit_op(Op::STR_INDEX_OF, 0);
        ci(c, 0);
        c.emit_op(Op::I32_GE_S, 0);
        c.emit_br_if(0, 0); // has 'e' → skip

        // while last char == '0' { RAW = RAW.slice(0, len-1) }
        {
            let zb = c.emit_block(0);
            let (zlp, _) = c.emit_loop_s(0);
            lg(c, RAW);
            c.emit_op(Op::STR_LENGTH, 0);
            c.emit_op(Op::I32_EQZ, 0);
            c.emit_br_if(1, 0);
            lg(c, RAW);
            lg(c, RAW);
            c.emit_op(Op::STR_LENGTH, 0);
            ci(c, 1);
            c.emit_op(Op::I32_SUB, 0);
            c.emit_op(Op::STR_CHAR_CODE_AT, 0);
            ci(c, 48);
            c.emit_op(Op::I32_NE, 0);
            c.emit_br_if(1, 0); // last != '0' → break
            lg(c, RAW);
            ci(c, 0);
            lg(c, RAW);
            c.emit_op(Op::STR_LENGTH, 0);
            ci(c, 1);
            c.emit_op(Op::I32_SUB, 0);
            c.emit_op(Op::STR_SLICE, 0);
            ls(c, RAW);
            c.emit_br(0, 0);
            c.emit_end(0);
            c.patch_loop(zlp);
            c.emit_end(0);
            c.patch_block(zb);
        }
        // drop a bare trailing '.'
        {
            let db = c.emit_block(0);
            lg(c, RAW);
            lg(c, RAW);
            c.emit_op(Op::STR_LENGTH, 0);
            ci(c, 1);
            c.emit_op(Op::I32_SUB, 0);
            c.emit_op(Op::STR_CHAR_CODE_AT, 0);
            ci(c, 46);
            c.emit_op(Op::I32_NE, 0);
            c.emit_br_if(0, 0); // last != '.' → skip
            lg(c, RAW);
            ci(c, 0);
            lg(c, RAW);
            c.emit_op(Op::STR_LENGTH, 0);
            ci(c, 1);
            c.emit_op(Op::I32_SUB, 0);
            c.emit_op(Op::STR_SLICE, 0);
            ls(c, RAW);
            c.emit_end(0);
            c.patch_block(db);
        }
        c.emit_end(0);
        c.patch_block(sb);
    }

    // %G → uppercase the exponent marker
    {
        let ub = c.emit_block(0);
        lg(c, CONV);
        ci(c, 71);
        c.emit_op(Op::I32_NE, 0);
        c.emit_br_if(0, 0);
        lg(c, RAW);
        hc(c, str_upper, 1);
        ls(c, RAW);
        c.emit_end(0);
        c.patch_block(ub);
    }

    sign_pos(c, str_cat);
    c.emit_end(0);
    c.patch_block(b);
}

fn conv_radix(
    c: &mut Chunk,
    math_trunc: u16,
    num_radix: u16,
    str_upper: u16,
    str_cat: u16,
    code: i32,
    radix: u16,
    uppercase: bool,
    prefix: &str,
) {
    let b = c.emit_block(0);
    lg(c, CONV);
    ci(c, code);
    c.emit_op(Op::I32_NE, 0);
    c.emit_br_if(0, 0);
    lg(c, N);
    hc(c, math_trunc, 1);
    cf(c, radix as f64);
    hc(c, num_radix, 2);
    if uppercase {
        hc(c, str_upper, 1);
    }
    ls(c, RAW);
    if !prefix.is_empty() {
        let ab = c.emit_block(0);
        lg(c, FALT);
        c.emit_op(Op::I32_EQZ, 0);
        c.emit_br_if(0, 0);
        lg(c, N);
        cf(c, 0.0);
        c.emit_op(Op::F64_EQ, 0);
        c.emit_br_if(0, 0);
        cs(c, prefix);
        lg(c, RAW);
        hc(c, str_cat, 2);
        ls(c, RAW);
        c.emit_end(0);
        c.patch_block(ab);
    }
    c.emit_end(0);
    c.patch_block(b);
}

fn conv_c(c: &mut Chunk, num_num: u16, str_fcc: u16) {
    let b = c.emit_block(0);
    lg(c, CONV);
    ci(c, 99);
    c.emit_op(Op::I32_NE, 0);
    c.emit_br_if(0, 0);
    lg(c, ARG);
    hc(c, num_num, 1);
    hc(c, str_fcc, 1);
    ls(c, RAW);
    c.emit_end(0);
    c.patch_block(b);
}

fn sign_pos(c: &mut Chunk, str_cat: u16) {
    let b = c.emit_block(0);
    lg(c, N);
    cf(c, 0.0);
    c.emit_op(Op::F64_LT, 0);
    c.emit_br_if(0, 0);
    {
        let sb = c.emit_block(0);
        lg(c, FSIGN);
        c.emit_op(Op::I32_EQZ, 0);
        c.emit_br_if(0, 0);
        cs(c, "+");
        lg(c, RAW);
        hc(c, str_cat, 2);
        ls(c, RAW);
        c.emit_br(1, 0);
        c.emit_end(0);
        c.patch_block(sb);
    }
    {
        let spb = c.emit_block(0);
        lg(c, FSPACE);
        c.emit_op(Op::I32_EQZ, 0);
        c.emit_br_if(0, 0);
        cs(c, " ");
        lg(c, RAW);
        hc(c, str_cat, 2);
        ls(c, RAW);
        c.emit_end(0);
        c.patch_block(spb);
    }
    c.emit_end(0);
    c.patch_block(b);
}

// ── sscanf helper ─────────────────────────────────────────────────────────
// Local slots for the `__fmt_sscanf` chunk.
const S_INP: u16 = 0; // input string (arg0)
const S_FMT: u16 = 1; // format string (arg1)
const S_I: u16 = 2; // input cursor
const S_J: u16 = 3; // format cursor
const S_ILEN: u16 = 4;
const S_FLEN: u16 = 5;
const S_OUT: u16 = 6; // result array
const S_FC: u16 = 7; // current format char code
const S_CONV: u16 = 8; // conversion char code
const S_IC: u16 = 9; // current input char code
const S_START: u16 = 10; // token start cursor
const S_NLOCALS: u16 = 11;

/// Push i32 1 when `S_IC` holds a whitespace char code.
fn ss_is_space(c: &mut Chunk) {
    lg(c, S_IC);
    ci(c, 32);
    c.emit_op(Op::I32_EQ, 0);
    lg(c, S_IC);
    ci(c, 9);
    c.emit_op(Op::I32_EQ, 0);
    c.emit_op(Op::I32_OR, 0);
    lg(c, S_IC);
    ci(c, 10);
    c.emit_op(Op::I32_EQ, 0);
    c.emit_op(Op::I32_OR, 0);
    lg(c, S_IC);
    ci(c, 13);
    c.emit_op(Op::I32_EQ, 0);
    c.emit_op(Op::I32_OR, 0);
}

/// Advance `S_I` over a run of whitespace in the input.
fn ss_skip_ws(c: &mut Chunk) {
    let wb = c.emit_block(0);
    let (wlp, _) = c.emit_loop_s(0);
    lg(c, S_I);
    lg(c, S_ILEN);
    c.emit_op(Op::I32_GE_S, 0);
    c.emit_br_if(1, 0);
    lg(c, S_INP);
    lg(c, S_I);
    c.emit_op(Op::STR_CHAR_CODE_AT, 0);
    ls(c, S_IC);
    ss_is_space(c);
    c.emit_op(Op::I32_EQZ, 0);
    c.emit_br_if(1, 0);
    inc(c, S_I);
    c.emit_br(0, 0);
    c.emit_end(0);
    c.patch_loop(wlp);
    c.emit_end(0);
    c.patch_block(wb);
}

/// Build the `__fmt_sscanf(input, fmt)` helper. Returns an array of the
/// values parsed out of `input` according to the C-style `fmt` (handles
/// `%d`/`%i`, `%f`, `%s`, `%c`, literal chars and whitespace runs).
pub fn build_sscanf(_imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new(SSCANF_CHUNK);
    c.arity = 2;
    c.local_count = S_NLOCALS;

    let num_num = c.add_import("ecma:number", "Number");
    let arr_push = c.add_import("ecma:array", "push");
    let arr_new = c.add_import("vybe:js-array", "newWithLength");

    // ilen / flen ; i = j = 0 ; out = []
    lg(&mut c, S_INP);
    c.emit_op(Op::STR_LENGTH, 0);
    ls(&mut c, S_ILEN);
    lg(&mut c, S_FMT);
    c.emit_op(Op::STR_LENGTH, 0);
    ls(&mut c, S_FLEN);
    c.emit_op(Op::I32_CONST_0, 0);
    ls(&mut c, S_I);
    c.emit_op(Op::I32_CONST_0, 0);
    ls(&mut c, S_J);
    c.emit_op(Op::I32_CONST_0, 0);
    hc(&mut c, arr_new, 1);
    ls(&mut c, S_OUT);

    let ob = c.emit_block(0);
    let (ol, _) = c.emit_loop_s(0);
    lg(&mut c, S_J);
    lg(&mut c, S_FLEN);
    c.emit_op(Op::I32_GE_S, 0);
    c.emit_br_if(1, 0);
    lg(&mut c, S_FMT);
    lg(&mut c, S_J);
    c.emit_op(Op::STR_CHAR_CODE_AT, 0);
    ls(&mut c, S_FC);

    // ── conversion specifier: '%' ─────────────────────────────────────────
    {
        let pct = c.emit_block(0);
        lg(&mut c, S_FC);
        ci(&mut c, 37);
        c.emit_op(Op::I32_NE, 0);
        c.emit_br_if(0, 0);
        inc(&mut c, S_J);
        lg(&mut c, S_FMT);
        lg(&mut c, S_J);
        c.emit_op(Op::STR_CHAR_CODE_AT, 0);
        ls(&mut c, S_CONV);
        inc(&mut c, S_J);
        ss_skip_ws(&mut c);

        // %d / %i / %f → parse a number token
        {
            let nb = c.emit_block(0);
            lg(&mut c, S_CONV);
            ci(&mut c, 100);
            c.emit_op(Op::I32_EQ, 0);
            lg(&mut c, S_CONV);
            ci(&mut c, 105);
            c.emit_op(Op::I32_EQ, 0);
            c.emit_op(Op::I32_OR, 0);
            lg(&mut c, S_CONV);
            ci(&mut c, 102);
            c.emit_op(Op::I32_EQ, 0);
            c.emit_op(Op::I32_OR, 0);
            c.emit_op(Op::I32_EQZ, 0);
            c.emit_br_if(0, 0);
            lg(&mut c, S_I);
            ls(&mut c, S_START);
            // optional leading sign
            {
                let sgn = c.emit_block(0);
                lg(&mut c, S_I);
                lg(&mut c, S_ILEN);
                c.emit_op(Op::I32_GE_S, 0);
                c.emit_br_if(0, 0);
                lg(&mut c, S_INP);
                lg(&mut c, S_I);
                c.emit_op(Op::STR_CHAR_CODE_AT, 0);
                ls(&mut c, S_IC);
                lg(&mut c, S_IC);
                ci(&mut c, 45);
                c.emit_op(Op::I32_EQ, 0);
                lg(&mut c, S_IC);
                ci(&mut c, 43);
                c.emit_op(Op::I32_EQ, 0);
                c.emit_op(Op::I32_OR, 0);
                c.emit_op(Op::I32_EQZ, 0);
                c.emit_br_if(0, 0);
                inc(&mut c, S_I);
                c.emit_end(0);
                c.patch_block(sgn);
            }
            // digit run (and '.' when parsing %f)
            {
                let dl = c.emit_block(0);
                let (dlp, _) = c.emit_loop_s(0);
                lg(&mut c, S_I);
                lg(&mut c, S_ILEN);
                c.emit_op(Op::I32_GE_S, 0);
                c.emit_br_if(1, 0);
                lg(&mut c, S_INP);
                lg(&mut c, S_I);
                c.emit_op(Op::STR_CHAR_CODE_AT, 0);
                ls(&mut c, S_IC);
                lg(&mut c, S_IC);
                ci(&mut c, 48);
                c.emit_op(Op::I32_GE_S, 0);
                lg(&mut c, S_IC);
                ci(&mut c, 57);
                c.emit_op(Op::I32_LE_S, 0);
                c.emit_op(Op::I32_AND, 0);
                lg(&mut c, S_IC);
                ci(&mut c, 46);
                c.emit_op(Op::I32_EQ, 0);
                lg(&mut c, S_CONV);
                ci(&mut c, 102);
                c.emit_op(Op::I32_EQ, 0);
                c.emit_op(Op::I32_AND, 0);
                c.emit_op(Op::I32_OR, 0);
                c.emit_op(Op::I32_EQZ, 0);
                c.emit_br_if(1, 0);
                inc(&mut c, S_I);
                c.emit_br(0, 0);
                c.emit_end(0);
                c.patch_loop(dlp);
                c.emit_end(0);
                c.patch_block(dl);
            }
            // out.push(Number(input.slice(start, i)))
            lg(&mut c, S_OUT);
            lg(&mut c, S_INP);
            lg(&mut c, S_START);
            lg(&mut c, S_I);
            c.emit_op(Op::STR_SLICE, 0);
            hc(&mut c, num_num, 1);
            hc(&mut c, arr_push, 2);
            c.emit_op(Op::DROP, 0);
            c.emit_br(2, 0); // continue ol
            c.emit_end(0);
            c.patch_block(nb);
        }

        // %s → non-whitespace token
        {
            let sb = c.emit_block(0);
            lg(&mut c, S_CONV);
            ci(&mut c, 115);
            c.emit_op(Op::I32_NE, 0);
            c.emit_br_if(0, 0);
            lg(&mut c, S_I);
            ls(&mut c, S_START);
            {
                let sl = c.emit_block(0);
                let (slp, _) = c.emit_loop_s(0);
                lg(&mut c, S_I);
                lg(&mut c, S_ILEN);
                c.emit_op(Op::I32_GE_S, 0);
                c.emit_br_if(1, 0);
                lg(&mut c, S_INP);
                lg(&mut c, S_I);
                c.emit_op(Op::STR_CHAR_CODE_AT, 0);
                ls(&mut c, S_IC);
                ss_is_space(&mut c);
                c.emit_br_if(1, 0);
                inc(&mut c, S_I);
                c.emit_br(0, 0);
                c.emit_end(0);
                c.patch_loop(slp);
                c.emit_end(0);
                c.patch_block(sl);
            }
            lg(&mut c, S_OUT);
            lg(&mut c, S_INP);
            lg(&mut c, S_START);
            lg(&mut c, S_I);
            c.emit_op(Op::STR_SLICE, 0);
            hc(&mut c, arr_push, 2);
            c.emit_op(Op::DROP, 0);
            c.emit_br(2, 0); // continue ol
            c.emit_end(0);
            c.patch_block(sb);
        }

        // %c → single character
        {
            let cb = c.emit_block(0);
            lg(&mut c, S_CONV);
            ci(&mut c, 99);
            c.emit_op(Op::I32_NE, 0);
            c.emit_br_if(0, 0);
            lg(&mut c, S_OUT);
            lg(&mut c, S_INP);
            lg(&mut c, S_I);
            c.emit_op(Op::STR_CHAR_AT, 0);
            hc(&mut c, arr_push, 2);
            c.emit_op(Op::DROP, 0);
            inc(&mut c, S_I);
            c.emit_br(2, 0); // continue ol
            c.emit_end(0);
            c.patch_block(cb);
        }

        c.emit_br(1, 0); // unknown conversion → continue ol
        c.emit_end(0);
        c.patch_block(pct);
    }

    // ── whitespace in format: match a run of whitespace in input ──────────
    {
        let wsb = c.emit_block(0);
        lg(&mut c, S_FC);
        ci(&mut c, 32);
        c.emit_op(Op::I32_EQ, 0);
        lg(&mut c, S_FC);
        ci(&mut c, 9);
        c.emit_op(Op::I32_EQ, 0);
        c.emit_op(Op::I32_OR, 0);
        lg(&mut c, S_FC);
        ci(&mut c, 10);
        c.emit_op(Op::I32_EQ, 0);
        c.emit_op(Op::I32_OR, 0);
        lg(&mut c, S_FC);
        ci(&mut c, 13);
        c.emit_op(Op::I32_EQ, 0);
        c.emit_op(Op::I32_OR, 0);
        c.emit_op(Op::I32_EQZ, 0);
        c.emit_br_if(0, 0);
        inc(&mut c, S_J);
        ss_skip_ws(&mut c);
        c.emit_br(1, 0); // continue ol
        c.emit_end(0);
        c.patch_block(wsb);
    }

    // ── literal char: must match the input, else stop scanning ────────────
    {
        let lit = c.emit_block(0);
        lg(&mut c, S_I);
        lg(&mut c, S_ILEN);
        c.emit_op(Op::I32_GE_S, 0);
        c.emit_br_if(2, 0); // input exhausted → break
        lg(&mut c, S_INP);
        lg(&mut c, S_I);
        c.emit_op(Op::STR_CHAR_CODE_AT, 0);
        lg(&mut c, S_FC);
        c.emit_op(Op::I32_NE, 0);
        c.emit_br_if(2, 0); // mismatch → break
        inc(&mut c, S_I);
        inc(&mut c, S_J);
        c.emit_end(0);
        c.patch_block(lit);
    }

    c.emit_br(0, 0); // continue ol
    c.emit_end(0);
    c.patch_loop(ol);
    c.emit_end(0);
    c.patch_block(ob);

    lg(&mut c, S_OUT);
    c.emit_op(Op::RETURN, 0);
    c
}
