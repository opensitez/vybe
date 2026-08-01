//! `csv` — delimiter-separated records as a shared primitive.
//!
//! CSV is a STRUCTURED format with a real grammar — quoting, doubled quotes,
//! embedded delimiters — so it belongs beside `json` and `url` rather than
//! inside one language's string adapter. RFC 4180 is the common core and every
//! language spells the same three knobs: delimiter, enclosure, escape.
//!
//! # Who needs this
//!
//! | language | surface | before |
//! |---|---|---|
//! | php | `str_getcsv`, `fputcsv` | 160-line scanner in `string_adapter.rs` |
//! | fortran | `str_getcsv` | DECLARED with no implementation — see below |
//! | pascal | `__pascal_file_readcsv`/`writecsv` | file I/O only, no parsing |
//! | python | `csv` module | declared as a known module |
//!
//! Go (`encoding/csv`) and Ruby (`CSV`) are the obvious next consumers.
//!
//! **fortran's `str_getcsv` could not run.** It resolved through
//! `LanguageHooks::str_getcsv`, which only php registers, reached via a
//! `profile.name == "fortran"` check in `builtins.rs` — a language-name test in
//! shared code — and then `.unwrap()` on `None`. Binding both languages here
//! removes the hook, the name check, and the panic together.
//!
//! # The dialect is RUNTIME, not compile-time
//!
//! php's `str_getcsv($s, $delimiter, $enclosure, $escape)` takes them as
//! arguments, so they are stack values. php's own emitter dropped them —
//! `for _ in 1..argc { DROP }` — so `str_getcsv($s, ';')` silently split on
//! commas. A compile-time option could not have served php at all.
//!
//! **No coercion here.** Callers pass strings.

use vybe_runtime::opcode::Op;
use vybe_runtime::{Chunk, Value};

fn get(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn set(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}

fn call(chunks: &mut [Chunk], current: usize, module: &str, name: &str, argc: u8, line: u32) {
    let idx = chunks[current].add_import(module.to_string(), name.to_string());
    chunks[current].emit_call(idx, argc, line);
}

/// One character of `s` at `idx`. Stack: `-> [string]`.
fn char_at(chunks: &mut [Chunk], current: usize, s: u16, idx: u16, line: u32) {
    get(&mut chunks[current], s, line);
    get(&mut chunks[current], idx, line);
    call(chunks, current, "ecma:string", "charAt", 2, line);
}

/// The default dialect: RFC 4180 — comma, double quote, quote-doubling.
/// A language pushes these when the caller supplied no explicit dialect.
pub fn emit_default_dialect(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_string_const(",", line);
    chunks[current].emit_string_const("\"", line);
}

/// Split one CSV record into its fields.
/// Stack: `[s, delimiter, enclosure]` → `[array]`.
///
/// A single-pass scanner with one piece of state — whether the cursor is inside
/// an enclosure. Inside, a doubled enclosure character is a literal one and
/// anything else is content, so an embedded delimiter survives; outside, the
/// enclosure opens a quoted field and the delimiter ends the current one.
///
/// The final field is always pushed, so `""` yields `[""]` and a trailing
/// delimiter yields a trailing empty field — which is what php, python and
/// RFC 4180 all specify.
pub fn emit_parse_line(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = chunks[current].alloc_scratch(8);
    let (enc, delim, s, out, cur, in_q, i, n) = (
        base,
        base + 1,
        base + 2,
        base + 3,
        base + 4,
        base + 5,
        base + 6,
        base + 7,
    );

    set(&mut chunks[current], enc, line);
    set(&mut chunks[current], delim, line);
    set(&mut chunks[current], s, line);

    crate::primitives::collections::emit_array_new(chunks, current, 0, line);
    set(&mut chunks[current], out, line);
    chunks[current].emit_string_const("", line);
    set(&mut chunks[current], cur, line);
    chunks[current].emit_bool_const(false, line);
    set(&mut chunks[current], in_q, line);
    chunks[current].emit_f64_const(0.0, line);
    set(&mut chunks[current], i, line);
    get(&mut chunks[current], s, line);
    call(chunks, current, "wasm:js-string", "length", 1, line);
    set(&mut chunks[current], n, line);

    let c = chunks[current].alloc_scratch(1);
    let loop_state = crate::primitives::loops::emit_loop_start(chunks, current, line);
    get(&mut chunks[current], i, line);
    get(&mut chunks[current], n, line);
    crate::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
    crate::primitives::loops::emit_loop_cond(chunks, current, line);

    char_at(chunks, current, s, i, line);
    set(&mut chunks[current], c, line);

    get(&mut chunks[current], in_q, line);
    crate::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    {
        // ── inside an enclosure ──
        get(&mut chunks[current], c, line);
        get(&mut chunks[current], enc, line);
        crate::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
        crate::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
        chunks[current].emit_if(line);
        {
            // A doubled enclosure is a literal one; anything else closes.
            let next = chunks[current].alloc_scratch(1);
            get(&mut chunks[current], i, line);
            chunks[current].emit_f64_const(1.0, line);
            chunks[current].emit_op(Op::F64_ADD, line);
            set(&mut chunks[current], next, line);

            get(&mut chunks[current], next, line);
            get(&mut chunks[current], n, line);
            crate::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
            crate::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
            chunks[current].emit_if(line);
            {
                char_at(chunks, current, s, next, line);
                get(&mut chunks[current], enc, line);
                crate::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
                crate::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
                chunks[current].emit_if(line);
                get(&mut chunks[current], cur, line);
                get(&mut chunks[current], enc, line);
                crate::primitives::ops::emit_dyn_add(&mut chunks[current], line);
                set(&mut chunks[current], cur, line);
                get(&mut chunks[current], next, line);
                set(&mut chunks[current], i, line);
                chunks[current].emit_else(line);
                chunks[current].emit_bool_const(false, line);
                set(&mut chunks[current], in_q, line);
                chunks[current].emit_end(line);
            }
            chunks[current].emit_else(line);
            chunks[current].emit_bool_const(false, line);
            set(&mut chunks[current], in_q, line);
            chunks[current].emit_end(line);
        }
        chunks[current].emit_else(line);
        get(&mut chunks[current], cur, line);
        get(&mut chunks[current], c, line);
        crate::primitives::ops::emit_dyn_add(&mut chunks[current], line);
        set(&mut chunks[current], cur, line);
        chunks[current].emit_end(line);
    }
    chunks[current].emit_else(line);
    {
        // ── outside an enclosure ──
        get(&mut chunks[current], c, line);
        get(&mut chunks[current], enc, line);
        crate::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
        crate::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
        chunks[current].emit_if(line);
        chunks[current].emit_bool_const(true, line);
        set(&mut chunks[current], in_q, line);
        chunks[current].emit_else(line);
        {
            get(&mut chunks[current], c, line);
            get(&mut chunks[current], delim, line);
            crate::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
            crate::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
            chunks[current].emit_if(line);
            get(&mut chunks[current], out, line);
            get(&mut chunks[current], cur, line);
            crate::primitives::collections::emit_push(chunks, current, line);
            chunks[current].emit_op(Op::DROP, line);
            chunks[current].emit_string_const("", line);
            set(&mut chunks[current], cur, line);
            chunks[current].emit_else(line);
            get(&mut chunks[current], cur, line);
            get(&mut chunks[current], c, line);
            crate::primitives::ops::emit_dyn_add(&mut chunks[current], line);
            set(&mut chunks[current], cur, line);
            chunks[current].emit_end(line);
        }
        chunks[current].emit_end(line);
    }
    chunks[current].emit_end(line);

    get(&mut chunks[current], i, line);
    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    set(&mut chunks[current], i, line);
    crate::primitives::loops::emit_loop_end(chunks, current, loop_state, line);

    get(&mut chunks[current], out, line);
    get(&mut chunks[current], cur, line);
    crate::primitives::collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    get(&mut chunks[current], out, line);
}

/// Render one record. Stack: `[array, delimiter, enclosure]` → `[string]`.
///
/// A field is enclosed only when it has to be — it contains the delimiter, the
/// enclosure, or a line break — which is what php `fputcsv` and python's
/// `csv.writer` with `QUOTE_MINIMAL` (the default) both do. An enclosure inside
/// the field is doubled.
pub fn emit_format_row(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = chunks[current].alloc_scratch(7);
    let (enc, delim, row, out, i, n, f) = (
        base,
        base + 1,
        base + 2,
        base + 3,
        base + 4,
        base + 5,
        base + 6,
    );

    set(&mut chunks[current], enc, line);
    set(&mut chunks[current], delim, line);
    set(&mut chunks[current], row, line);

    chunks[current].emit_string_const("", line);
    set(&mut chunks[current], out, line);
    chunks[current].emit_f64_const(0.0, line);
    set(&mut chunks[current], i, line);
    get(&mut chunks[current], row, line);
    chunks[current].emit_op(Op::ARRAY_LENGTH, line);
    set(&mut chunks[current], n, line);

    let loop_state = crate::primitives::loops::emit_loop_start(chunks, current, line);
    get(&mut chunks[current], i, line);
    get(&mut chunks[current], n, line);
    crate::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
    crate::primitives::loops::emit_loop_cond(chunks, current, line);

    // A separator before every field but the first.
    get(&mut chunks[current], i, line);
    chunks[current].emit_f64_const(0.0, line);
    crate::primitives::ops::emit_dyn_gt(&mut chunks[current], line);
    crate::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    get(&mut chunks[current], out, line);
    get(&mut chunks[current], delim, line);
    crate::primitives::ops::emit_dyn_add(&mut chunks[current], line);
    set(&mut chunks[current], out, line);
    chunks[current].emit_end(line);

    get(&mut chunks[current], row, line);
    get(&mut chunks[current], i, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    call(chunks, current, "ecma:string", "String", 1, line);
    set(&mut chunks[current], f, line);

    // needs_quotes = f contains delim, enc, "\n" or "\r"
    let needs = chunks[current].alloc_scratch(1);
    chunks[current].emit_i32_const(0, line);
    set(&mut chunks[current], needs, line);
    for probe in [delim, enc] {
        get(&mut chunks[current], f, line);
        get(&mut chunks[current], probe, line);
        call(chunks, current, "ecma:string", "includes", 2, line);
        crate::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
        get(&mut chunks[current], needs, line);
        chunks[current].emit_op(Op::I32_OR, line);
        set(&mut chunks[current], needs, line);
    }
    for nl in ["\n", "\r"] {
        get(&mut chunks[current], f, line);
        chunks[current].emit_string_const(nl, line);
        call(chunks, current, "ecma:string", "includes", 2, line);
        crate::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
        get(&mut chunks[current], needs, line);
        chunks[current].emit_op(Op::I32_OR, line);
        set(&mut chunks[current], needs, line);
    }

    get(&mut chunks[current], needs, line);
    chunks[current].emit_if_value(line);
    {
        // enc ++ f.replaceAll(enc, enc+enc) ++ enc
        get(&mut chunks[current], enc, line);
        get(&mut chunks[current], f, line);
        get(&mut chunks[current], enc, line);
        get(&mut chunks[current], enc, line);
        get(&mut chunks[current], enc, line);
        crate::primitives::ops::emit_dyn_add(&mut chunks[current], line);
        call(chunks, current, "ecma:string", "replaceAll", 3, line);
        crate::primitives::ops::emit_dyn_add(&mut chunks[current], line);
        get(&mut chunks[current], enc, line);
        crate::primitives::ops::emit_dyn_add(&mut chunks[current], line);
    }
    chunks[current].emit_else(line);
    get(&mut chunks[current], f, line);
    chunks[current].emit_end(line);
    let rendered = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], rendered, line);

    get(&mut chunks[current], out, line);
    get(&mut chunks[current], rendered, line);
    crate::primitives::ops::emit_dyn_add(&mut chunks[current], line);
    set(&mut chunks[current], out, line);

    get(&mut chunks[current], i, line);
    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    set(&mut chunks[current], i, line);
    crate::primitives::loops::emit_loop_end(chunks, current, loop_state, line);

    get(&mut chunks[current], out, line);
}

/// Silence the unused-import warning for `Value` when only some emitters are
/// compiled in; kept because the module builds constants in both directions.
#[allow(dead_code)]
fn _value_marker(_v: Value) {}
