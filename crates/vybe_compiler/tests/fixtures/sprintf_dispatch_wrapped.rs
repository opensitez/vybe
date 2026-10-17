#![allow(dead_code)]
use vybe_compiler::primitives::instructions::core_wasm;
use vybe_runtime::{Chunk, Op};
const CHUNK_NAME: &str = "__fmt_sprintf";
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
}
fn ci(c: &mut Chunk, v: i32) {
    c.emit_i32_const(v, 0);
}
fn cf(c: &mut Chunk, v: f64) {
    c.emit_f64_const(v, 0);
}
fn cs(c: &mut Chunk, v: &str) {
    c.emit_string_const(v, 0);
}
fn hc(c: &mut Chunk, i: u16, a: u8) {
    c.emit_call(i, a, 0);
}
fn inc(c: &mut Chunk, s: u16) {
    lg(c, s);
    core_wasm::i32_const(c, 0, 1);
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
    let str_index = c.add_import("ecma:string", "indexOf");
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

    // Exact %s needs only argument conversion, without scanning or copying
    // the format. Extra arguments are ignored by the general path too.
    lg(&mut c, FMT);
    let str_test = c.add_import("wasm:js-string", "test");
    hc(&mut c, str_test, 1);
    c.emit_if(0);
    lg(&mut c, FMT);
    cs(&mut c, "%s");
    let str_equals = c.add_import("wasm:js-string", "equals");
    hc(&mut c, str_equals, 2);
    c.emit_if(0);
    lg(&mut c, ARGS);
    ci(&mut c, 0);
    hc(&mut c, arr_at, 2);
    hc(&mut c, str_tostr, 1);
    c.emit_op(Op::RETURN, 0);
    c.emit_end(0);
    c.emit_end(0);

    // init
    lg(&mut c, FMT);
    hc(&mut c, str_len, 1);
    ls(&mut c, FLEN);

    cs(&mut c, "");
    ls(&mut c, OUT);
    core_wasm::i32_const(&mut c, 0, 0);
    ls(&mut c, I);
    core_wasm::i32_const(&mut c, 0, 0);
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
        lg(&mut c, CH);
        ci(&mut c, 37);
        c.emit_op(Op::I32_NE, 0);
        c.emit_if(0);
        // Copy the complete literal run in one operation, preserving UTF-16
        // surrogate pairs and avoiding repeated copies of the growing output.
        lg(&mut c, OUT);
        lg(&mut c, FMT);
        lg(&mut c, I);
        lg(&mut c, FMT);
        cs(&mut c, "%");
        lg(&mut c, I);
        hc(&mut c, str_index, 3);
        c.emit_op(Op::I32_FROM_F64, 0);
        c.emit_op_u16(Op::LOCAL_TEE, I, 0);
        ci(&mut c, 0);
        c.emit_op(Op::I32_LT_S, 0);
        c.emit_if(0);
        lg(&mut c, FLEN);
        ls(&mut c, I);
        c.emit_end(0);
        lg(&mut c, I);
        hc(&mut c, str_slice, 3);
        hc(&mut c, str_cat, 2);
        ls(&mut c, OUT);
        c.emit_br(1, 0); // continue ol
        c.emit_end(0);
    }
    inc(&mut c, I); // skip '%'

    // Plain %%/%s/%c skip the general flag/width/precision parser.
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

        // Plain %s -> next argument as a string. This handles every plain
        // string conversion without the separate single-%s scan and recipe.
        {
            let string = c.emit_block(0);
            lg(&mut c, CH);
            ci(&mut c, 115);
            c.emit_op(Op::I32_NE, 0);
            c.emit_br_if(0, 0);
            lg(&mut c, OUT);
            lg(&mut c, ARGS);
            lg(&mut c, AIDX);
            hc(&mut c, arr_at, 2);
            hc(&mut c, str_tostr, 1);
            hc(&mut c, str_cat, 2);
            ls(&mut c, OUT);
            inc(&mut c, AIDX);
            inc(&mut c, I);
            c.emit_br(2, 0);
            c.emit_end(0);
            c.patch_block(string);
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
            {
                let idx = c.add_import("wasm:js-string", "test");
                c.emit_call(idx, 1, 0);
            }
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
        core_wasm::i32_const(&mut c, 0, 0);
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
    core_wasm::i32_const(&mut c, 0, 0);
    ls(&mut c, FLEFT);
    core_wasm::i32_const(&mut c, 0, 0);
    ls(&mut c, FSIGN);
    core_wasm::i32_const(&mut c, 0, 0);
    ls(&mut c, FZERO);
    core_wasm::i32_const(&mut c, 0, 0);
    ls(&mut c, FSPACE);
    core_wasm::i32_const(&mut c, 0, 0);
    ls(&mut c, FALT);
    core_wasm::i32_const(&mut c, 0, 0);
    ls(&mut c, WIDTH);
    core_wasm::i32_const(&mut c, 0, 0);
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
        core_wasm::i32_const(&mut c, 0, 0);
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
    lg(&mut c, CH);
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

    // A table selects exactly the former conversion case; argument Number
    // conversion and RAW initialization still occur before dispatch.
    conv_s(&mut c, str_tostr, str_slice);
    let done = c.emit_block(0);
    let radix_join = c.emit_block(0);
    let mut labels = [0usize; 11];
    for index in (1..11).rev() {
        labels[index] = c.emit_block(0);
    }
    let mut table = [11u32; 52];
    for (code, case) in [
        (100, 1),
        (105, 1),
        (117, 2),
        (102, 3),
        (70, 3),
        (101, 4),
        (69, 4),
        (120, 5),
        (88, 6),
        (111, 7),
        (98, 8),
        (103, 9),
        (71, 9),
        (99, 10),
    ] {
        table[(code - 69) as usize] = case - 1;
    }
    lg(&mut c, CONV);
    ci(&mut c, 69);
    c.emit_op(Op::I32_SUB, 0);
    c.emit_br_table(&table, 11, 0);
    for case in 1..11 {
        c.emit_end(0);
        c.patch_block(labels[case]);
        let remaining = 10 - case;
        if (5..=8).contains(&case) {
            let (radix, prefix, upper) = match case {
                5 => (16.0, "0x", 0),
                6 => (16.0, "0X", 1),
                7 => (8.0, "0", 0),
                _ => (2.0, "", 0),
            };
            cf(&mut c, radix);
            ls(&mut c, POS);
            cs(&mut c, prefix);
            ls(&mut c, CH);
            ci(&mut c, upper);
            ls(&mut c, SAVEI);
            c.emit_br(remaining as u32, 0);
        } else {
            let start = c.code.len();
            match case {
                1 => conv_d(&mut c, math_trunc, math_abs, num_radix, str_cat),
                2 => conv_u(&mut c, num_radix),
                3 => conv_f(&mut c, num_fixed, str_cat, math_pow),
                4 => conv_e(&mut c, num_exp, str_upper, str_cat),
                9 => conv_g(&mut c, num_prec, str_upper, str_cat),
                10 => conv_c(&mut c, num_num, str_fcc),
                _ => unreachable!(),
            }
            audit_case_branches(&c, start);
            c.emit_br((remaining + 1) as u32, 0);
        }
    }
    c.emit_end(0);
    c.patch_block(radix_join);
    // POS/CH/SAVEI parser work slots are dead after ARG has been selected.
    lg(&mut c, N);
    hc(&mut c, math_trunc, 1);
    lg(&mut c, POS);
    hc(&mut c, num_radix, 2);
    ls(&mut c, RAW);
    lg(&mut c, SAVEI);
    c.emit_if(0);
    lg(&mut c, RAW);
    hc(&mut c, str_upper, 1);
    ls(&mut c, RAW);
    c.emit_end(0);
    let prefix = c.emit_block(0);
    lg(&mut c, FALT);
    c.emit_op(Op::I32_EQZ, 0);
    c.emit_br_if(0, 0);
    lg(&mut c, N);
    cf(&mut c, 0.0);
    c.emit_op(Op::F64_EQ, 0);
    c.emit_br_if(0, 0);
    lg(&mut c, CONV);
    ci(&mut c, 98);
    c.emit_op(Op::I32_EQ, 0);
    c.emit_br_if(0, 0);
    lg(&mut c, CH);
    lg(&mut c, RAW);
    hc(&mut c, str_cat, 2);
    ls(&mut c, RAW);
    c.emit_end(0);
    c.patch_block(prefix);
    c.emit_end(0);
    c.patch_block(done);
    c.emit_end(0); // original string case

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
        core_wasm::i32_const(&mut c, 0, 0);
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
            {
                let idx = c.add_import("wasm:js-string", "charCodeAt");
                c.emit_call(idx, 2, 0);
            }
            ci(&mut c, 48);
            c.emit_op(Op::I32_NE, 0);
            c.emit_br_if(0, 0);
            lg(&mut c, RAW);
            {
                let idx = c.add_import("wasm:js-string", "length");
                c.emit_call(idx, 1, 0);
            }
            c.emit_op(Op::I32_EQZ, 0);
            c.emit_br_if(0, 0);
            lg(&mut c, RAW);
            ci(&mut c, 0);
            {
                let idx = c.add_import("wasm:js-string", "charCodeAt");
                c.emit_call(idx, 2, 0);
            }
            ci(&mut c, 43);
            c.emit_op(Op::I32_EQ, 0);
            lg(&mut c, RAW);
            ci(&mut c, 0);
            {
                let idx = c.add_import("wasm:js-string", "charCodeAt");
                c.emit_call(idx, 2, 0);
            }
            ci(&mut c, 45);
            c.emit_op(Op::I32_EQ, 0);
            c.emit_op(Op::I32_OR, 0);
            lg(&mut c, RAW);
            ci(&mut c, 0);
            {
                let idx = c.add_import("wasm:js-string", "charCodeAt");
                c.emit_call(idx, 2, 0);
            }
            ci(&mut c, 32);
            c.emit_op(Op::I32_EQ, 0);
            c.emit_op(Op::I32_OR, 0);
            c.emit_op(Op::I32_EQZ, 0);
            c.emit_br_if(0, 0);
            lg(&mut c, RAW);
            ci(&mut c, 0);
            {
                let idx = c.add_import("ecma:string", "charAt");
                c.emit_call(idx, 2, 0);
            }
            ls(&mut c, CUSTPAD);
            lg(&mut c, CUSTPAD);
            lg(&mut c, RAW);
            ci(&mut c, 1);
            lg(&mut c, RAW);
            {
                let idx = c.add_import("wasm:js-string", "length");
                c.emit_call(idx, 1, 0);
            }
            {
                let idx = c.add_import("ecma:string", "slice");
                c.emit_call(idx, 3, 0);
            }
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
    lg(c, CONV);
    ci(c, 115);
    c.emit_op(Op::I32_EQ, 0);
    c.emit_if(0);
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
    c.emit_else(0);
}

fn conv_d(c: &mut Chunk, math_trunc: u16, math_abs: u16, num_radix: u16, str_cat: u16) {
    c.emit_block(0);
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
}

fn conv_u(c: &mut Chunk, num_radix: u16) {
    c.emit_block(0);
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
}

fn conv_f(c: &mut Chunk, num_fixed: u16, str_cat: u16, math_pow: u16) {
    c.emit_block(0);

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
    {
        let idx = c.add_import("wasm:js-string", "length");
        c.emit_call(idx, 1, 0);
    }
    ci(c, 2);
    c.emit_op(Op::I32_LT_S, 0);
    c.emit_br_if(0, 0);
    lg(c, RAW);
    ci(c, 0);
    {
        let idx = c.add_import("wasm:js-string", "charCodeAt");
        c.emit_call(idx, 2, 0);
    }
    ci(c, 45);
    c.emit_op(Op::I32_NE, 0);
    c.emit_br_if(0, 0);
    lg(c, RAW);
    ci(c, 1);
    {
        let idx = c.add_import("wasm:js-string", "charCodeAt");
        c.emit_call(idx, 2, 0);
    }
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
    {
        let idx = c.add_import("wasm:js-string", "length");
        c.emit_call(idx, 1, 0);
    }
    {
        let idx = c.add_import("ecma:string", "slice");
        c.emit_call(idx, 3, 0);
    }
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
    c.emit_block(0);
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
}

// %g / %G — shortest of %e/%f at `precision` significant digits, with
// trailing zeros (and a bare trailing '.') removed. Uses Number.toPrecision
// then strips, which matches PHP for the common fixed-notation cases.
fn conv_g(c: &mut Chunk, num_prec: u16, str_upper: u16, str_cat: u16) {
    c.emit_block(0);

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
        {
            let idx = c.add_import("ecma:string", "indexOf");
            c.emit_call(idx, 2, 0);
        }
        ci(c, 0);
        c.emit_op(Op::I32_LT_S, 0);
        c.emit_br_if(0, 0); // no '.' → skip
        lg(c, RAW);
        cs(c, "e");
        {
            let idx = c.add_import("ecma:string", "indexOf");
            c.emit_call(idx, 2, 0);
        }
        ci(c, 0);
        c.emit_op(Op::I32_GE_S, 0);
        c.emit_br_if(0, 0); // has 'e' → skip

        // while last char == '0' { RAW = RAW.slice(0, len-1) }
        {
            let zb = c.emit_block(0);
            let (zlp, _) = c.emit_loop_s(0);
            lg(c, RAW);
            {
                let idx = c.add_import("wasm:js-string", "length");
                c.emit_call(idx, 1, 0);
            }
            c.emit_op(Op::I32_EQZ, 0);
            c.emit_br_if(1, 0);
            lg(c, RAW);
            lg(c, RAW);
            {
                let idx = c.add_import("wasm:js-string", "length");
                c.emit_call(idx, 1, 0);
            }
            ci(c, 1);
            c.emit_op(Op::I32_SUB, 0);
            {
                let idx = c.add_import("wasm:js-string", "charCodeAt");
                c.emit_call(idx, 2, 0);
            }
            ci(c, 48);
            c.emit_op(Op::I32_NE, 0);
            c.emit_br_if(1, 0); // last != '0' → break
            lg(c, RAW);
            ci(c, 0);
            lg(c, RAW);
            {
                let idx = c.add_import("wasm:js-string", "length");
                c.emit_call(idx, 1, 0);
            }
            ci(c, 1);
            c.emit_op(Op::I32_SUB, 0);
            {
                let idx = c.add_import("ecma:string", "slice");
                c.emit_call(idx, 3, 0);
            }
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
            {
                let idx = c.add_import("wasm:js-string", "length");
                c.emit_call(idx, 1, 0);
            }
            ci(c, 1);
            c.emit_op(Op::I32_SUB, 0);
            {
                let idx = c.add_import("wasm:js-string", "charCodeAt");
                c.emit_call(idx, 2, 0);
            }
            ci(c, 46);
            c.emit_op(Op::I32_NE, 0);
            c.emit_br_if(0, 0); // last != '.' → skip
            lg(c, RAW);
            ci(c, 0);
            lg(c, RAW);
            {
                let idx = c.add_import("wasm:js-string", "length");
                c.emit_call(idx, 1, 0);
            }
            ci(c, 1);
            c.emit_op(Op::I32_SUB, 0);
            {
                let idx = c.add_import("ecma:string", "slice");
                c.emit_call(idx, 3, 0);
            }
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
    lg(c, CONV);
    ci(c, code);
    c.emit_op(Op::I32_EQ, 0);
    c.emit_if(0);
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
    c.emit_else(0);
}

fn conv_c(c: &mut Chunk, num_num: u16, str_fcc: u16) {
    c.emit_block(0);
    lg(c, ARG);
    hc(c, num_num, 1);
    hc(c, str_fcc, 1);
    ls(c, RAW);
    c.emit_end(0);
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

// Each old arm's outer IF became a block. A branch may reach that block,
// but must never depend on one of the former enclosing dispatch IFs.
fn audit_case_branches(c: &Chunk, start: usize) {
    let mut ip = start;
    let mut depth = 0u32;
    while ip < c.code.len() {
        let op = Op::from_code_at(&c.code, ip);
        if matches!(op, Op::BLOCK | Op::LOOP | Op::IF) {
            depth += 1;
        }
        if matches!(op, Op::BR | Op::BR_IF) {
            let mut at = ip + 4;
            let target = vybe_runtime::opcode::read_leb_u32(&c.code, &mut at);
            assert!(
                target < depth,
                "case branch escapes its own original arm: target={target} depth={depth}"
            );
        }
        if op == Op::END {
            depth -= 1;
        }
        ip += 4 + op.operand_format().size_in(&c.code, ip + 4);
    }
    assert_eq!(depth, 0);
}
