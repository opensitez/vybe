//! .NET date/time PARSING, driven by the same pattern language the formatter
//! prints.
//!
//! One helper chunk consumes a pattern and an input together. `Parse` tries the
//! invariant patterns .NET accepts, in order, and falls back to
//! `ecma:date.parse` for the ISO forms the host already handles.

use vybe_runtime::opcode::Op;
use vybe_runtime::Chunk;

use super::datetime_adapter;

/// Month abbreviations, 3 chars each — `indexOf` gives the month directly.
/// Read from the FORMATTER's table so the two cannot drift on spelling.
const MONTH_ABBR: &str = super::datetime_format_adapter::MONTH_ABBR;

/// The patterns `Parse` tries, longest/most-specific first. Every one of them
/// is a format .NET's invariant and en-US cultures accept.
///
/// ⛔ Order matters: `"M/d/yyyy"` consumes the date half of
/// `"5/14/2024 3:45:59 PM"` and then fails on the leftover, so the forms WITH a
/// time come first. The helper only succeeds on WHOLE-input consumption, which
/// is what makes trying them in sequence safe.
const PARSE_PATTERNS: &[&str] = &[
    "yyyy-MM-ddTHH:mm:ss.fff",
    "yyyy-MM-ddTHH:mm:ss",
    "yyyy-MM-dd HH:mm:ss",
    "yyyy-MM-dd HH:mm",
    "yyyy-MM-dd",
    "M/d/yyyy h:mm:ss tt",
    "M/d/yyyy h:mm tt",
    "M/d/yyyy HH:mm:ss",
    "M/d/yyyy HH:mm",
    "M/d/yyyy",
    "MMM d, yyyy",
    "d MMM yyyy",
];

/// TIME-ONLY forms, deliberately NOT in [`PARSE_PATTERNS`].
///
/// ⛔ A bare time is not a `DateTime.Parse` input: .NET gives it TODAY's date.
/// Matching it in the sweep would land on year 1. `TimeValue` is the function
/// whose job IS a bare time, and year 1 is what it wants, so they live there.
const TIME_PATTERNS: &[&str] = &["h:mm:ss tt", "h:mm tt", "HH:mm:ss", "HH:mm"];

fn push_str(chunk: &mut Chunk, text: &str, line: u32) {
    chunk.emit_string_const(text, line);
}

fn get(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn set(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}

fn call(chunk: &mut Chunk, module: &str, name: &str, argc: u8, line: u32) {
    let idx = chunk.add_import(module, name);
    chunk.emit_call(idx, argc, line);
}

fn str_len(chunk: &mut Chunk, line: u32) {
    call(chunk, "ecma:string", "length", 1, line);
}

fn char_at(chunk: &mut Chunk, line: u32) {
    call(chunk, "ecma:string", "charAt", 2, line);
}

fn char_code_at(chunk: &mut Chunk, line: u32) {
    call(chunk, "ecma:string", "charCodeAt", 2, line);
}

fn substring(chunk: &mut Chunk, line: u32) {
    call(chunk, "wasm:js-string", "substring", 3, line);
}

fn str_eq(chunk: &mut Chunk, line: u32) {
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
}

struct Slots {
    i: u16,
    j: u16,
    flen: u16,
    ilen: u16,
    ch: u16,
    run: u16,
    ok: u16,
    num: u16,
    taken: u16,
    tmp: u16,
    year: u16,
    month: u16,
    day: u16,
    hour: u16,
    minute: u16,
    second: u16,
    millis: u16,
    pm: u16,
    code: u16,
    lit: u16,
    offset: u16,
}

fn alloc_slots(chunk: &mut Chunk) -> Slots {
    let b = chunk.alloc_scratch(21);
    Slots {
        i: b,
        j: b + 1,
        flen: b + 2,
        ilen: b + 3,
        ch: b + 4,
        run: b + 5,
        ok: b + 6,
        num: b + 7,
        taken: b + 8,
        tmp: b + 9,
        year: b + 10,
        month: b + 11,
        day: b + 12,
        hour: b + 13,
        minute: b + 14,
        second: b + 15,
        millis: b + 16,
        pm: b + 17,
        code: b + 18,
        lit: b + 19,
        offset: b + 20,
    }
}

/// Read up to `max` decimal digits at `j` into `num`, advancing `j`.
///
/// `min` digits are REQUIRED — `"dd"` must see two, `"d"` accepts one or two,
/// which is what makes `M/d/yyyy` match both `5/14/2024` and `12/3/2024`.
fn emit_read_digits(chunk: &mut Chunk, s: &Slots, min: i32, max: i32, line: u32) {
    chunk.emit_f64_const(0.0, line);
    set(chunk, s.num, line);
    chunk.emit_f64_const(0.0, line);
    set(chunk, s.taken, line);

    let block = chunk.emit_block(line);
    let (lp, _) = chunk.emit_loop_s(line);
    // taken == max → stop
    get(chunk, s.taken, line);
    chunk.emit_f64_const(f64::from(max), line);
    str_eq(chunk, line);
    chunk.emit_br_if(1, line);
    // j >= ilen → stop
    get(chunk, s.j, line);
    get(chunk, s.ilen, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_not(chunk, line);
    chunk.emit_br_if(1, line);
    // code = input.charCodeAt(j); not a digit → stop
    get(chunk, 0, line);
    get(chunk, s.j, line);
    char_code_at(chunk, line);
    set(chunk, s.code, line);
    get(chunk, s.code, line);
    chunk.emit_f64_const(48.0, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_br_if(1, line);
    chunk.emit_f64_const(57.0, line);
    get(chunk, s.code, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_br_if(1, line);

    get(chunk, s.num, line);
    chunk.emit_f64_const(10.0, line);
    chunk.emit_op(Op::F64_MUL, line);
    get(chunk, s.code, line);
    chunk.emit_f64_const(48.0, line);
    chunk.emit_op(Op::F64_SUB, line);
    chunk.emit_op(Op::F64_ADD, line);
    set(chunk, s.num, line);
    get(chunk, s.j, line);
    chunk.emit_f64_const(1.0, line);
    chunk.emit_op(Op::F64_ADD, line);
    set(chunk, s.j, line);
    get(chunk, s.taken, line);
    chunk.emit_f64_const(1.0, line);
    chunk.emit_op(Op::F64_ADD, line);
    set(chunk, s.taken, line);
    chunk.emit_br(0, line);
    chunk.emit_end(line);
    chunk.patch_loop(lp);
    chunk.emit_end(line);
    chunk.patch_block(block);

    // Too few digits → the parse failed.
    get(chunk, s.taken, line);
    chunk.emit_f64_const(f64::from(min), line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    chunk.emit_i32_const(0, line);
    set(chunk, s.ok, line);
    chunk.emit_end(line);
}

/// Read EXACTLY `run` digits — the count the pattern asked for.
///
/// The fixed-width `emit_read_digits` cannot express it: `run` is only known
/// while the loop is running.
fn emit_read_run_digits(chunk: &mut Chunk, s: &Slots, line: u32) {
    chunk.emit_f64_const(0.0, line);
    set(chunk, s.num, line);
    chunk.emit_f64_const(0.0, line);
    set(chunk, s.taken, line);

    let block = chunk.emit_block(line);
    let (lp, _) = chunk.emit_loop_s(line);
    get(chunk, s.taken, line);
    get(chunk, s.run, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_not(chunk, line);
    chunk.emit_br_if(1, line);
    get(chunk, s.j, line);
    get(chunk, s.ilen, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_not(chunk, line);
    chunk.emit_br_if(1, line);
    get(chunk, 0, line);
    get(chunk, s.j, line);
    char_code_at(chunk, line);
    set(chunk, s.code, line);
    get(chunk, s.code, line);
    chunk.emit_f64_const(48.0, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_br_if(1, line);
    chunk.emit_f64_const(57.0, line);
    get(chunk, s.code, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_br_if(1, line);
    get(chunk, s.num, line);
    chunk.emit_f64_const(10.0, line);
    chunk.emit_op(Op::F64_MUL, line);
    get(chunk, s.code, line);
    chunk.emit_f64_const(48.0, line);
    chunk.emit_op(Op::F64_SUB, line);
    chunk.emit_op(Op::F64_ADD, line);
    set(chunk, s.num, line);
    get(chunk, s.j, line);
    chunk.emit_f64_const(1.0, line);
    chunk.emit_op(Op::F64_ADD, line);
    set(chunk, s.j, line);
    get(chunk, s.taken, line);
    chunk.emit_f64_const(1.0, line);
    chunk.emit_op(Op::F64_ADD, line);
    set(chunk, s.taken, line);
    chunk.emit_br(0, line);
    chunk.emit_end(line);
    chunk.patch_loop(lp);
    chunk.emit_end(line);
    chunk.patch_block(block);

    get(chunk, s.taken, line);
    get(chunk, s.run, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    chunk.emit_i32_const(0, line);
    set(chunk, s.ok, line);
    chunk.emit_end(line);
}

/// `if ok && ch == c { … }` — the same flat ladder shape the formatter uses,
/// with the running `ok` folded into the guard so a failed field short-circuits
/// the rest.
fn arm<F: FnOnce(&mut Chunk)>(chunk: &mut Chunk, s: &Slots, c: &str, line: u32, body: F) {
    get(chunk, s.tmp, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_if(line);
    get(chunk, s.ch, line);
    push_str(chunk, c, line);
    str_eq(chunk, line);
    chunk.emit_if(line);
    body(chunk);
    chunk.emit_i32_const(1, line);
    set(chunk, s.tmp, line);
    chunk.emit_end(line);
    chunk.emit_end(line);
}

/// A numeric field: read `run >= 2 ? 2 : 1` minimum digits into `slot`.
fn numeric_arm(chunk: &mut Chunk, s: &Slots, slot: u16, max: i32, line: u32) {
    get(chunk, s.run, line);
    chunk.emit_f64_const(2.0, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_not(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    emit_read_digits(chunk, s, 2, max, line);
    chunk.emit_else(line);
    emit_read_digits(chunk, s, 1, max, line);
    chunk.emit_end(line);
    get(chunk, s.num, line);
    set(chunk, slot, line);
}

/// Match the one-character literal in `slot` at the input cursor, or fail.
fn emit_match_literal_slot(chunk: &mut Chunk, s: &Slots, slot: u16, line: u32) {
    get(chunk, 0, line);
    get(chunk, s.j, line);
    char_at(chunk, line);
    get(chunk, slot, line);
    str_eq(chunk, line);
    chunk.emit_if(line);
    get(chunk, s.j, line);
    chunk.emit_f64_const(1.0, line);
    chunk.emit_op(Op::F64_ADD, line);
    set(chunk, s.j, line);
    chunk.emit_else(line);
    chunk.emit_i32_const(0, line);
    set(chunk, s.ok, line);
    chunk.emit_end(line);
}

/// `__dotnet_date_parse_exact(input, fmt) → ms | NaN`
fn build_chunk(line: u32) -> Chunk {
    let mut m = Chunk::new("__dotnet_date_parse_exact");
    m.arity = 2;
    m.local_count = 2;
    let s = alloc_slots(&mut m);
    let c = &mut m;

    for (slot, init) in [
        (s.year, 1.0),
        (s.month, 1.0),
        (s.day, 1.0),
        (s.hour, 0.0),
        (s.minute, 0.0),
        (s.second, 0.0),
        (s.millis, 0.0),
        (s.pm, -1.0),
        (s.i, 0.0),
        (s.j, 0.0),
    ] {
        c.emit_f64_const(init, line);
        set(c, slot, line);
    }
    c.emit_f64_const(0.0, line);
    set(c, s.offset, line);
    c.emit_i32_const(1, line);
    set(c, s.ok, line);

    // ⛔ A ONE-CHARACTER format is a STANDARD specifier, the same rule the
    // formatter follows — `ParseExact(s, "o", …)` asks for the round-trip
    // pattern, not for a literal `o`. The table is the formatter's, so the two
    // cannot disagree about what `"o"` means.
    get(c, 1, line);
    set(c, s.lit, line);
    get(c, s.lit, line);
    str_len(c, line);
    c.emit_f64_const(1.0, line);
    c.emit_op(Op::F64_EQ, line);
    c.emit_if(line);
    for (spec, pattern) in super::datetime_format_adapter::STANDARD_PATTERNS {
        get(c, s.lit, line);
        push_str(c, spec, line);
        str_eq(c, line);
        c.emit_if(line);
        push_str(c, pattern, line);
        set(c, s.lit, line);
        c.emit_end(line);
    }
    c.emit_end(line);

    get(c, s.lit, line);
    str_len(c, line);
    set(c, s.flen, line);
    get(c, 0, line);
    str_len(c, line);
    set(c, s.ilen, line);

    let block_p = c.emit_block(line);
    let (loop_p, _) = c.emit_loop_s(line);

    get(c, s.ok, line);
    c.emit_op(Op::I32_EQZ, line);
    c.emit_br_if(1, line);
    get(c, s.i, line);
    get(c, s.flen, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(c, line);
    vybe_compiler::primitives::ops::emit_dyn_not(c, line);
    c.emit_br_if(1, line);

    get(c, s.lit, line);
    get(c, s.i, line);
    char_at(c, line);
    set(c, s.ch, line);

    // run length of the repeat
    c.emit_f64_const(1.0, line);
    set(c, s.run, line);
    let ib = c.emit_block(line);
    let (il, _) = c.emit_loop_s(line);
    get(c, s.i, line);
    get(c, s.run, line);
    c.emit_op(Op::F64_ADD, line);
    get(c, s.flen, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(c, line);
    vybe_compiler::primitives::ops::emit_dyn_not(c, line);
    c.emit_br_if(1, line);
    get(c, s.lit, line);
    get(c, s.i, line);
    get(c, s.run, line);
    c.emit_op(Op::F64_ADD, line);
    char_at(c, line);
    get(c, s.ch, line);
    str_eq(c, line);
    c.emit_op(Op::I32_EQZ, line);
    c.emit_br_if(1, line);
    get(c, s.run, line);
    c.emit_f64_const(1.0, line);
    c.emit_op(Op::F64_ADD, line);
    set(c, s.run, line);
    c.emit_br(0, line);
    c.emit_end(line);
    c.patch_loop(il);
    c.emit_end(line);
    c.patch_block(ib);

    c.emit_i32_const(0, line);
    set(c, s.tmp, line);

    arm(c, &s, "y", line, |c| {
        get(c, s.run, line);
        c.emit_f64_const(4.0, line);
        vybe_compiler::primitives::ops::emit_dyn_lt(c, line);
        vybe_compiler::primitives::ops::emit_dyn_not(c, line);
        vybe_compiler::primitives::ops::emit_dyn_to_bool(c, line);
        c.emit_if(line);
        emit_read_digits(c, &s, 4, 4, line);
        get(c, s.num, line);
        set(c, s.year, line);
        c.emit_else(line);
        emit_read_digits(c, &s, 2, 2, line);
        // Two-digit years are 2000-based below 30, 1900-based at or above —
        // .NET's invariant `TwoDigitYearMax` of 2029.
        get(c, s.num, line);
        c.emit_f64_const(30.0, line);
        vybe_compiler::primitives::ops::emit_dyn_lt(c, line);
        vybe_compiler::primitives::ops::emit_dyn_to_bool(c, line);
        c.emit_if_value(line);
        c.emit_f64_const(2000.0, line);
        c.emit_else(line);
        c.emit_f64_const(1900.0, line);
        c.emit_end(line);
        get(c, s.num, line);
        c.emit_op(Op::F64_ADD, line);
        set(c, s.year, line);
        c.emit_end(line);
    });

    arm(c, &s, "M", line, |c| {
        get(c, s.run, line);
        c.emit_f64_const(3.0, line);
        vybe_compiler::primitives::ops::emit_dyn_lt(c, line);
        vybe_compiler::primitives::ops::emit_dyn_not(c, line);
        vybe_compiler::primitives::ops::emit_dyn_to_bool(c, line);
        c.emit_if(line);
        // A name: locate the 3-letter abbreviation in the table.
        push_str(c, MONTH_ABBR, line);
        get(c, 0, line);
        get(c, s.j, line);
        get(c, s.j, line);
        c.emit_f64_const(3.0, line);
        c.emit_op(Op::F64_ADD, line);
        substring(c, line);
        c.emit_f64_const(0.0, line);
        call(c, "ecma:string", "indexOf", 3, line);
        set(c, s.num, line);
        get(c, s.num, line);
        c.emit_f64_const(0.0, line);
        vybe_compiler::primitives::ops::emit_dyn_lt(c, line);
        vybe_compiler::primitives::ops::emit_dyn_to_bool(c, line);
        c.emit_if(line);
        c.emit_i32_const(0, line);
        set(c, s.ok, line);
        c.emit_else(line);
        get(c, s.num, line);
        c.emit_f64_const(3.0, line);
        c.emit_op(Op::F64_DIV, line);
        c.emit_f64_const(1.0, line);
        c.emit_op(Op::F64_ADD, line);
        set(c, s.month, line);
        get(c, s.j, line);
        c.emit_f64_const(3.0, line);
        c.emit_op(Op::F64_ADD, line);
        set(c, s.j, line);
        c.emit_end(line);
        c.emit_else(line);
        numeric_arm(c, &s, s.month, 2, line);
        c.emit_end(line);
    });

    arm(c, &s, "d", line, |c| numeric_arm(c, &s, s.day, 2, line));
    arm(c, &s, "H", line, |c| numeric_arm(c, &s, s.hour, 2, line));
    arm(c, &s, "h", line, |c| numeric_arm(c, &s, s.hour, 2, line));
    arm(c, &s, "m", line, |c| numeric_arm(c, &s, s.minute, 2, line));
    arm(c, &s, "s", line, |c| numeric_arm(c, &s, s.second, 2, line));
    // `f`/`F` — EXACTLY `run` digits, scaled onto milliseconds. `"o"` asks for
    // seven, and the whole-input rule rejects any left unconsumed.
    for spec in ["f", "F"] {
        arm(c, &s, spec, line, |c| {
            emit_read_run_digits(c, &s, line);
            get(c, s.num, line);
            c.emit_f64_const(10.0, line);
            c.emit_f64_const(3.0, line);
            get(c, s.run, line);
            c.emit_op(Op::F64_SUB, line);
            call(c, "ecma:math", "pow", 2, line);
            c.emit_op(Op::F64_MUL, line);
            c.emit_op(Op::F64_TRUNC, line);
            set(c, s.millis, line);
        });
    }

    // `K` — `Z`, a signed `+HH:mm`, or nothing at all.
    arm(c, &s, "K", line, |c| {
        get(c, 0, line);
        get(c, s.j, line);
        get(c, s.j, line);
        c.emit_f64_const(1.0, line);
        c.emit_op(Op::F64_ADD, line);
        substring(c, line);
        set(c, s.tmp2(), line);
        get(c, s.tmp2(), line);
        push_str(c, "Z", line);
        str_eq(c, line);
        c.emit_if(line);
        get(c, s.j, line);
        c.emit_f64_const(1.0, line);
        c.emit_op(Op::F64_ADD, line);
        set(c, s.j, line);
        c.emit_end(line);
        for (sign_text, sign) in [("+", 1.0f64), ("-", -1.0f64)] {
            get(c, s.tmp2(), line);
            push_str(c, sign_text, line);
            str_eq(c, line);
            c.emit_if(line);
            get(c, s.j, line);
            c.emit_f64_const(1.0, line);
            c.emit_op(Op::F64_ADD, line);
            set(c, s.j, line);
            emit_read_digits(c, &s, 2, 2, line);
            get(c, s.num, line);
            c.emit_f64_const(3_600_000.0 * sign, line);
            c.emit_op(Op::F64_MUL, line);
            set(c, s.offset, line);
            push_str(c, ":", line);
            set(c, s.tmp2(), line);
            emit_match_literal_slot(c, &s, s.tmp2(), line);
            emit_read_digits(c, &s, 2, 2, line);
            get(c, s.offset, line);
            get(c, s.num, line);
            c.emit_f64_const(60_000.0 * sign, line);
            c.emit_op(Op::F64_MUL, line);
            c.emit_op(Op::F64_ADD, line);
            set(c, s.offset, line);
            c.emit_end(line);
        }
    });

    arm(c, &s, "t", line, |c| {
        get(c, 0, line);
        get(c, s.j, line);
        get(c, s.j, line);
        c.emit_f64_const(1.0, line);
        c.emit_op(Op::F64_ADD, line);
        substring(c, line);
        call(c, "ecma:string", "toUpperCase", 1, line);
        set(c, s.lit, line);
        get(c, s.lit, line);
        push_str(c, "P", line);
        str_eq(c, line);
        c.emit_if_value(line);
        c.emit_f64_const(1.0, line);
        c.emit_else(line);
        get(c, s.lit, line);
        push_str(c, "A", line);
        str_eq(c, line);
        c.emit_if_value(line);
        c.emit_f64_const(0.0, line);
        c.emit_else(line);
        c.emit_i32_const(0, line);
        set(c, s.ok, line);
        c.emit_f64_const(-1.0, line);
        c.emit_end(line);
        c.emit_end(line);
        set(c, s.pm, line);
        // `tt` consumes the `M` as well.
        get(c, s.j, line);
        get(c, s.run, line);
        c.emit_op(Op::F64_ADD, line);
        set(c, s.j, line);
    });

    // `\c` — the next pattern character, matched literally.
    arm(c, &s, "\\", line, |c| {
        c.emit_f64_const(2.0, line);
        set(c, s.run, line);
        get(c, s.lit, line);
        get(c, s.i, line);
        c.emit_f64_const(1.0, line);
        c.emit_op(Op::F64_ADD, line);
        char_at(c, line);
        set(c, s.tmp2(), line);
        emit_match_literal_slot(c, &s, s.tmp2(), line);
    });

    // Anything else: the pattern character itself, once per repeat.
    get(c, s.tmp, line);
    c.emit_op(Op::I32_EQZ, line);
    c.emit_if(line);
    get(c, s.ch, line);
    set(c, s.tmp2(), line);
    c.emit_f64_const(0.0, line);
    set(c, s.taken, line);
    let lb = c.emit_block(line);
    let (ll, _) = c.emit_loop_s(line);
    get(c, s.taken, line);
    get(c, s.run, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(c, line);
    vybe_compiler::primitives::ops::emit_dyn_not(c, line);
    c.emit_br_if(1, line);
    emit_match_literal_slot(c, &s, s.tmp2(), line);
    get(c, s.taken, line);
    c.emit_f64_const(1.0, line);
    c.emit_op(Op::F64_ADD, line);
    set(c, s.taken, line);
    c.emit_br(0, line);
    c.emit_end(line);
    c.patch_loop(ll);
    c.emit_end(line);
    c.patch_block(lb);
    c.emit_end(line);

    get(c, s.i, line);
    get(c, s.run, line);
    c.emit_op(Op::F64_ADD, line);
    set(c, s.i, line);
    c.emit_br(0, line);
    c.emit_end(line);
    c.patch_loop(loop_p);
    c.emit_end(line);
    c.patch_block(block_p);

    // ⛔ The WHOLE input must be consumed. That is what lets `Parse` try a list
    // of patterns in sequence: `"M/d/yyyy"` matches the front of
    // `"5/14/2024 3:45:59 PM"` and is rejected here for the leftover.
    get(c, s.j, line);
    get(c, s.ilen, line);
    str_eq(c, line);
    c.emit_op(Op::I32_EQZ, line);
    c.emit_if(line);
    c.emit_i32_const(0, line);
    set(c, s.ok, line);
    c.emit_end(line);

    // AM/PM folds onto the 24-hour clock.
    get(c, s.pm, line);
    c.emit_f64_const(1.0, line);
    str_eq(c, line);
    c.emit_if(line);
    get(c, s.hour, line);
    c.emit_f64_const(12.0, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(c, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(c, line);
    c.emit_if(line);
    get(c, s.hour, line);
    c.emit_f64_const(12.0, line);
    c.emit_op(Op::F64_ADD, line);
    set(c, s.hour, line);
    c.emit_end(line);
    c.emit_end(line);
    get(c, s.pm, line);
    c.emit_f64_const(0.0, line);
    str_eq(c, line);
    c.emit_if(line);
    get(c, s.hour, line);
    c.emit_f64_const(12.0, line);
    str_eq(c, line);
    c.emit_if(line);
    c.emit_f64_const(0.0, line);
    set(c, s.hour, line);
    c.emit_end(line);
    c.emit_end(line);

    get(c, s.ok, line);
    c.emit_if_value(line);
    get(c, s.year, line);
    get(c, s.month, line);
    c.emit_f64_const(1.0, line);
    c.emit_op(Op::F64_SUB, line);
    get(c, s.day, line);
    get(c, s.hour, line);
    get(c, s.minute, line);
    get(c, s.second, line);
    get(c, s.millis, line);
    call(c, "ecma:date", "UTC", 7, line);
    // A `K`/`zzz` offset means the reading was LOCAL to that offset.
    get(c, s.offset, line);
    c.emit_op(Op::F64_SUB, line);
    c.emit_else(line);
    c.emit_f64_const(f64::NAN, line);
    c.emit_end(line);
    c.emit_op(Op::RETURN, line);
    m
}

impl Slots {
    /// A second scratch string. `lit` holds the (possibly EXPANDED) pattern
    /// for the whole run, so a one-character literal needs its own slot.
    fn tmp2(&self) -> u16 {
        self.code
    }
}

/// `__dotnet_date_parse_exact_any(input, fmt) → ms | NaN`.
///
/// `fmt` may be a single pattern OR an ARRAY of them — .NET's
/// `ParseExact(s, String(), provider, styles)` overload, which
/// `test_vb_date_time_parse_exact_multiple_formats_array` uses. The first
/// pattern that consumes the whole input wins.
fn build_any_chunk(exact_idx: usize, line: u32) -> Chunk {
    let mut m = Chunk::new("__dotnet_date_parse_exact_any");
    m.arity = 2;
    m.local_count = 2;
    let result = m.alloc_scratch(4);
    let fnref = result + 1;
    let k = result + 2;
    let n = result + 3;
    let c = &mut m;

    c.emit_op_u16(Op::REF_FUNC, exact_idx as u16, line);
    c.emit(0, line);
    set(c, fnref, line);

    // A string is one pattern; anything else is the array overload.
    get(c, 1, line);
    call(c, "ecma:value", "typeof", 1, line);
    push_str(c, "string", line);
    str_eq(c, line);
    c.emit_if_value(line);
    get(c, fnref, line);
    get(c, 0, line);
    get(c, 1, line);
    c.emit_op_u8_u8(Op::CALL_REF, 2, 1, line);
    c.emit_else(line);
    c.emit_f64_const(f64::NAN, line);
    set(c, result, line);
    c.emit_f64_const(0.0, line);
    set(c, k, line);
    get(c, 1, line);
    call(c, "ecma:array", "length", 1, line);
    set(c, n, line);
    let block = c.emit_block(line);
    let (lp, _) = c.emit_loop_s(line);
    get(c, k, line);
    get(c, n, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(c, line);
    vybe_compiler::primitives::ops::emit_dyn_not(c, line);
    c.emit_br_if(1, line);
    get(c, result, line);
    get(c, result, line);
    c.emit_op(Op::F64_EQ, line);
    c.emit_br_if(1, line);
    get(c, fnref, line);
    get(c, 0, line);
    get(c, 1, line);
    get(c, k, line);
    c.emit_op(Op::ARRAY_GET, line);
    c.emit_op_u8_u8(Op::CALL_REF, 2, 1, line);
    set(c, result, line);
    get(c, k, line);
    c.emit_f64_const(1.0, line);
    c.emit_op(Op::F64_ADD, line);
    set(c, k, line);
    c.emit_br(0, line);
    c.emit_end(line);
    c.patch_loop(lp);
    c.emit_end(line);
    c.patch_block(block);
    get(c, result, line);
    c.emit_end(line);
    c.emit_op(Op::RETURN, line);
    m
}

fn any_chunk(chunks: &mut Vec<Chunk>, line: u32) -> usize {
    if let Some(idx) = chunks
        .iter()
        .position(|c| c.name == "__dotnet_date_parse_exact_any")
    {
        return idx;
    }
    let exact = parse_chunk(chunks, line);
    chunks.push(build_any_chunk(exact, line));
    chunks.len() - 1
}

fn parse_chunk(chunks: &mut Vec<Chunk>, line: u32) -> usize {
    if let Some(idx) = chunks
        .iter()
        .position(|c| c.name == "__dotnet_date_parse_exact")
    {
        return idx;
    }
    chunks.push(build_chunk(line));
    chunks.len() - 1
}

/// `[input, fmt] → [ms | NaN]`
///
/// ⛔ Takes its scratch from the caller. It is emitted once per candidate
/// PATTERN, and a fresh `alloc_scratch(3)` each time would grow the enclosing
/// function by three locals per pattern — in a compiler whose scratch slots
/// alias named locals.
fn emit_call_parse_into(
    chunks: &mut Vec<Chunk>,
    current: usize,
    input_slot: u16,
    fmt_slot: u16,
    fn_slot: u16,
    line: u32,
) {
    let idx = any_chunk(chunks, line);
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_SET, fmt_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, input_slot, line);
    chunk.emit_op_u16(Op::REF_FUNC, idx as u16, line);
    chunk.emit(0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, fn_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, fn_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, input_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, fmt_slot, line);
    chunk.emit_op_u8_u8(Op::CALL_REF, 2, 1, line);
}

fn emit_call_parse(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let base = chunks[current].alloc_scratch(3);
    emit_call_parse_into(chunks, current, base, base + 1, base + 2, line);
}

/// `__dotnet_date_parse(input) → ms | NaN` — the whole sweep as ONE chunk.
///
/// ⛔ A chunk, not an inline expansion: twenty patterns is ~1,100 instructions
/// and `DateTime.Parse` appears everywhere. The sweep is a pure function of its
/// input, so one copy per module serves all call sites.
fn build_sweep_chunk(exact_idx: usize, line: u32) -> Chunk {
    let mut m = Chunk::new("__dotnet_date_parse");
    m.arity = 1;
    m.local_count = 1;
    let result = m.alloc_scratch(4);
    let fnref = result + 1;
    let tod = result + 2;
    let today = result + 3;
    let c = &mut m;

    c.emit_op_u16(Op::REF_FUNC, exact_idx as u16, line);
    c.emit(0, line);
    set(c, fnref, line);
    c.emit_f64_const(f64::NAN, line);
    set(c, result, line);

    // NaN is the only value not equal to itself, so `r != r` IS "no pattern has
    // matched yet". Each candidate runs only while that holds.
    let mut sweep = |c: &mut Chunk, pattern: &str| {
        get(c, result, line);
        get(c, result, line);
        c.emit_op(Op::F64_NE, line);
        c.emit_if(line);
        get(c, fnref, line);
        get(c, 0, line);
        push_str(c, pattern, line);
        c.emit_op_u8_u8(Op::CALL_REF, 2, 1, line);
        set(c, result, line);
        c.emit_end(line);
    };
    for pattern in PARSE_PATTERNS {
        sweep(c, pattern);
    }

    // Still nothing: the host's parser, which knows RFC 3339/2822 and the
    // long-form month spellings.
    get(c, result, line);
    get(c, result, line);
    c.emit_op(Op::F64_NE, line);
    c.emit_if(line);
    get(c, 0, line);
    call(c, "ecma:date", "parse", 1, line);
    set(c, result, line);
    c.emit_end(line);

    // Last: a bare TIME, which .NET puts on TODAY's date. Kept out of the
    // sweep above, where the year would default to 1.
    get(c, result, line);
    get(c, result, line);
    c.emit_op(Op::F64_NE, line);
    c.emit_if(line);
    c.emit_f64_const(f64::NAN, line);
    set(c, tod, line);
    for pattern in TIME_PATTERNS {
        get(c, tod, line);
        get(c, tod, line);
        c.emit_op(Op::F64_NE, line);
        c.emit_if(line);
        get(c, fnref, line);
        get(c, 0, line);
        push_str(c, pattern, line);
        c.emit_op_u8_u8(Op::CALL_REF, 2, 1, line);
        set(c, tod, line);
        c.emit_end(line);
    }
    get(c, tod, line);
    get(c, tod, line);
    c.emit_op(Op::F64_EQ, line);
    c.emit_if(line);
    call(c, "ecma:date", "now", 0, line);
    c.emit_f64_const(86_400_000.0, line);
    c.emit_op(Op::F64_DIV, line);
    c.emit_op(Op::F64_FLOOR, line);
    c.emit_f64_const(86_400_000.0, line);
    c.emit_op(Op::F64_MUL, line);
    set(c, today, line);
    get(c, today, line);
    emit_time_of_day(c, tod, line);
    c.emit_op(Op::F64_ADD, line);
    set(c, result, line);
    c.emit_end(line);
    c.emit_end(line);

    get(c, result, line);
    c.emit_op(Op::RETURN, line);
    m
}

/// `[] → [ms within the day]` for the instant in `ms_slot`, always positive.
fn emit_time_of_day(chunk: &mut Chunk, ms_slot: u16, line: u32) {
    get(chunk, ms_slot, line);
    chunk.emit_f64_const(86_400_000.0, line);
    vybe_compiler::primitives::math::emit_c_fmod(chunk, line);
    chunk.emit_f64_const(86_400_000.0, line);
    chunk.emit_op(Op::F64_ADD, line);
    chunk.emit_f64_const(86_400_000.0, line);
    vybe_compiler::primitives::math::emit_c_fmod(chunk, line);
}

fn sweep_chunk(chunks: &mut Vec<Chunk>, line: u32) -> usize {
    if let Some(idx) = chunks.iter().position(|c| c.name == "__dotnet_date_parse") {
        return idx;
    }
    let exact = parse_chunk(chunks, line);
    chunks.push(build_sweep_chunk(exact, line));
    chunks.len() - 1
}

/// `[input] → [ms | NaN]`
fn emit_parse_millis(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let idx = sweep_chunk(chunks, line);
    let chunk = &mut chunks[current];
    let input_slot = chunk.alloc_scratch(2);
    let fn_slot = input_slot + 1;
    chunk.emit_op_u16(Op::LOCAL_SET, input_slot, line);
    chunk.emit_op_u16(Op::REF_FUNC, idx as u16, line);
    chunk.emit(0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, fn_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, fn_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, input_slot, line);
    chunk.emit_op_u8_u8(Op::CALL_REF, 1, 1, line);
}

/// `DateTime.Parse(s)`.
pub fn emit_datetime_parse(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    emit_parse_millis(chunks, current, line);
    datetime_adapter::emit_datetime_from_millis(chunks, current, line);
}

/// `DateTime.Parse(s)` for `TryParse` — NULL rather than a value when `s` is
/// not a date.
pub fn emit_datetime_try_parse(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    emit_parse_millis(chunks, current, line);
    let chunk = &mut chunks[current];
    let ms_slot = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, ms_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, ms_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, ms_slot, line);
    chunk.emit_op(Op::F64_NE, line);
    chunk.emit_if_value(line);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, ms_slot, line);
    datetime_adapter::emit_datetime_from_millis(chunks, current, line);
    chunks[current].emit_end(line);
}

/// `DateTime.ParseExact(s, fmt, …)`.
pub fn emit_datetime_parse_exact(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    {
        let chunk = &mut chunks[current];
        // A provider / styles argument past the format sits on top.
        for _ in 2..argc {
            chunk.emit_op(Op::DROP, line);
        }
    }
    if argc < 2 {
        emit_parse_millis(chunks, current, line);
    } else {
        emit_call_parse(chunks, current, line);
    }
    datetime_adapter::emit_datetime_from_millis(chunks, current, line);
}

/// `TryParseExact` — NULL when the input does not match the format.
pub fn emit_datetime_try_parse_exact(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    {
        let chunk = &mut chunks[current];
        for _ in 2..argc {
            chunk.emit_op(Op::DROP, line);
        }
    }
    if argc < 2 {
        emit_parse_millis(chunks, current, line);
    } else {
        emit_call_parse(chunks, current, line);
    }
    let chunk = &mut chunks[current];
    let ms_slot = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, ms_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, ms_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, ms_slot, line);
    chunk.emit_op(Op::F64_NE, line);
    chunk.emit_if_value(line);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, ms_slot, line);
    datetime_adapter::emit_datetime_from_millis(chunks, current, line);
    chunks[current].emit_end(line);
}

/// `TimeValue(s)` — the TIME of day alone, on `DateTime.MinValue`'s date.
/// Real VB answers `#1/1/0001 3:45:00 PM#` for `"3:45 PM"`, which is what
/// makes `CStr` print the time without a date.
pub fn emit_vb_timevalue(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    emit_parse_millis(chunks, current, line);
    let chunk = &mut chunks[current];
    let ms_slot = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, ms_slot, line);
    // ⛔ NOT `Date.UTC(1, 0, 1)` for the date half: ECMA-262 §21.4.3.4 maps a
    // year in `[0, 99]` onto `1900 + y`, so that call answers 1901. The epoch
    // offset of `DateTime.MinValue` is the constant for year 1.
    chunk.emit_f64_const(
        vybe_compiler::primitives::datetime::DOTNET_DATETIME_MIN_UNIX_MS,
        line,
    );
    emit_time_of_day(chunk, ms_slot, line);
    chunk.emit_op(Op::F64_ADD, line);
    datetime_adapter::emit_datetime_from_millis(chunks, current, line);
}

/// `DateValue(s)` — the DATE alone, time zeroed.
pub fn emit_vb_datevalue(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    emit_parse_millis(chunks, current, line);
    let chunk = &mut chunks[current];
    chunk.emit_f64_const(86_400_000.0, line);
    chunk.emit_op(Op::F64_DIV, line);
    chunk.emit_op(Op::F64_FLOOR, line);
    chunk.emit_f64_const(86_400_000.0, line);
    chunk.emit_op(Op::F64_MUL, line);
    datetime_adapter::emit_datetime_from_millis(chunks, current, line);
}
