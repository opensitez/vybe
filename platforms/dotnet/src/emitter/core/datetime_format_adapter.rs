//! .NET custom & standard date/time format strings — one runtime helper chunk
//! that walks the pattern, so any format string works.
//!
//! Scope, from the split `primitives/datetime.rs` documents: the FORMAT
//! LETTERS are .NET spelling and belong in the adapter. Everything the letters
//! stand for is read off the value the shared emitters already built.
//!
//! ⛔ `DateTimeOffset` shares this: its object carries the same `Year`/`Month`/
//! … fields plus `__offset_ms`, so `zzz` answers from the value rather than
//! from a second implementation.

use std::sync::Arc;
use vybe_runtime::opcode::Op;
use vybe_runtime::{Chunk, Value};

/// Month names, FIXED WIDTH so a name is one `substring` away — no array
/// construction and no per-month branch. 9 = `"September"`.
const MONTH_NAMES: &str = "January  February March    April    May      June     \
July     August   SeptemberOctober  November December ";
pub(crate) const MONTH_ABBR: &str = "JanFebMarAprMayJunJulAugSepOctNovDec";
/// Day names, same trick. 9 = `"Wednesday"`.
const DAY_NAMES: &str = "Sunday   Monday   Tuesday  WednesdayThursday Friday   Saturday ";
const DAY_ABBR: &str = "SunMonTueWedThuFriSat";
const NAME_WIDTH: i32 = 9;

/// The STANDARD format specifiers, each as the custom pattern .NET defines it
/// to mean under the invariant culture. A one-character format string is a
/// STANDARD specifier — `ToString("d")` is the short date, not the day of the
/// month; `"%d"` is how you ask for the single custom specifier.
pub(crate) const STANDARD_PATTERNS: &[(&str, &str)] = &[
    ("d", "M/d/yyyy"),
    ("D", "dddd, MMMM d, yyyy"),
    ("t", "h:mm tt"),
    ("T", "h:mm:ss tt"),
    ("f", "dddd, MMMM d, yyyy h:mm tt"),
    ("F", "dddd, MMMM d, yyyy h:mm:ss tt"),
    ("g", "M/d/yyyy h:mm tt"),
    ("G", "M/d/yyyy h:mm:ss tt"),
    ("m", "MMMM d"),
    ("M", "MMMM d"),
    ("y", "MMMM yyyy"),
    ("Y", "MMMM yyyy"),
    ("s", "yyyy-MM-ddTHH:mm:ss"),
    ("u", "yyyy-MM-dd HH:mm:ssZ"),
    ("o", "yyyy-MM-ddTHH:mm:ss.fffffffK"),
    ("O", "yyyy-MM-ddTHH:mm:ss.fffffffK"),
    ("r", "ddd, dd MMM yyyy HH:mm:ss GMT"),
    ("R", "ddd, dd MMM yyyy HH:mm:ss GMT"),
];

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

/// `[s] → [n]`
fn str_len(chunk: &mut Chunk, line: u32) {
    call(chunk, "ecma:string", "length", 1, line);
}

/// `[s, i] → [char]`
fn char_at(chunk: &mut Chunk, line: u32) {
    call(chunk, "ecma:string", "charAt", 2, line);
}

/// `[s, a, b] → [sub]`
fn substring(chunk: &mut Chunk, line: u32) {
    call(chunk, "wasm:js-string", "substring", 3, line);
}

/// `[a, b] → [ab]`
fn concat(chunk: &mut Chunk, line: u32) {
    call(chunk, "wasm:js-string", "concat", 2, line);
}

/// `[v] → [str]`
fn to_str(chunk: &mut Chunk, line: u32) {
    call(chunk, "ecma:string", "String", 1, line);
}

/// `[n] → [str]`, zero-padded on the left to `width`.
fn pad(chunk: &mut Chunk, width: i32, line: u32) {
    to_str(chunk, line);
    chunk.emit_i32_const(width, line);
    push_str(chunk, "0", line);
    call(chunk, "ecma:string", "padStart", 3, line);
}

/// A field of the DateTime / DateTimeOffset in local 0. `[] → [value]`
fn field(chunk: &mut Chunk, name: &str, line: u32) {
    get(chunk, 0, line);
    let key = chunk.add_constant(Value::String(Arc::from(name)));
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, key, line);
}

/// `[a, b] → [i32 0/1]`
fn str_eq(chunk: &mut Chunk, line: u32) {
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
}

/// Slot layout of the helper chunk. Named so the emit below reads like the
/// algorithm rather than like arithmetic on a base index.
struct Slots {
    i: u16,
    len: u16,
    out: u16,
    ch: u16,
    run: u16,
    piece: u16,
    matched: u16,
    tmp: u16,
    year: u16,
    month: u16,
    day: u16,
    hour: u16,
    minute: u16,
    second: u16,
    millis: u16,
    dow: u16,
    offset_ms: u16,
    hour12: u16,
    off_abs: u16,
    fmt: u16,
    has_offset: u16,
}

fn alloc_slots(chunk: &mut Chunk) -> Slots {
    let b = chunk.alloc_scratch(21);
    Slots {
        i: b,
        len: b + 1,
        out: b + 2,
        ch: b + 3,
        run: b + 4,
        piece: b + 5,
        matched: b + 6,
        tmp: b + 7,
        year: b + 8,
        month: b + 9,
        day: b + 10,
        hour: b + 11,
        minute: b + 12,
        second: b + 13,
        millis: b + 14,
        dow: b + 15,
        offset_ms: b + 16,
        hour12: b + 17,
        off_abs: b + 18,
        fmt: b + 19,
        has_offset: b + 20,
    }
}

/// One arm of the specifier ladder: `if !matched && ch == c { … ; matched = 1 }`.
///
/// ⛔ Flat, not a nested if/else chain: fifteen hand-placed `end`s in bytecode
/// is a shape where one misplacement silently reparents the rest.
fn arm<F: FnOnce(&mut Chunk)>(chunk: &mut Chunk, s: &Slots, c: &str, line: u32, body: F) {
    get(chunk, s.matched, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_if(line);
    get(chunk, s.ch, line);
    push_str(chunk, c, line);
    str_eq(chunk, line);
    chunk.emit_if(line);
    body(chunk);
    set(chunk, s.piece, line);
    chunk.emit_i32_const(1, line);
    set(chunk, s.matched, line);
    chunk.emit_end(line);
    chunk.emit_end(line);
}

/// `[] → [str]` — the fixed-width name table entry for `index_slot`.
fn name_at(chunk: &mut Chunk, table: &str, index_slot: u16, width: i32, trim: bool, line: u32) {
    push_str(chunk, table, line);
    get(chunk, index_slot, line);
    chunk.emit_f64_const(f64::from(width), line);
    chunk.emit_op(Op::F64_MUL, line);
    get(chunk, index_slot, line);
    chunk.emit_f64_const(f64::from(width), line);
    chunk.emit_op(Op::F64_MUL, line);
    chunk.emit_f64_const(f64::from(width), line);
    chunk.emit_op(Op::F64_ADD, line);
    substring(chunk, line);
    if trim {
        call(chunk, "ecma:string", "trim", 1, line);
    }
}

/// `+HH:mm` / `-HH:mm` — the signed offset, the spelling `zzz` and `K` share.
/// `[] → [str]`
fn emit_offset_text(chunk: &mut Chunk, s: &Slots, line: u32) {
    get(chunk, s.offset_ms, line);
    chunk.emit_f64_const(0.0, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    push_str(chunk, "-", line);
    chunk.emit_else(line);
    push_str(chunk, "+", line);
    chunk.emit_end(line);
    get(chunk, s.off_abs, line);
    chunk.emit_f64_const(60.0, line);
    chunk.emit_op(Op::F64_DIV, line);
    chunk.emit_op(Op::F64_TRUNC, line);
    pad(chunk, 2, line);
    concat(chunk, line);
    push_str(chunk, ":", line);
    concat(chunk, line);
    get(chunk, s.off_abs, line);
    chunk.emit_f64_const(60.0, line);
    vybe_compiler::primitives::math::emit_c_fmod(chunk, line);
    pad(chunk, 2, line);
    concat(chunk, line);
}

/// `run >= n ? <a> : <b>` — `[] → [str]`
fn run_ge<A: FnOnce(&mut Chunk), B: FnOnce(&mut Chunk)>(
    chunk: &mut Chunk,
    s: &Slots,
    n: i32,
    line: u32,
    a: A,
    b: B,
) {
    get(chunk, s.run, line);
    chunk.emit_f64_const(f64::from(n), line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_not(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    a(chunk);
    chunk.emit_else(line);
    b(chunk);
    chunk.emit_end(line);
}

/// `__dotnet_date_format(value, fmt)` — the interpreter.
fn build_chunk(line: u32) -> Chunk {
    let mut m = Chunk::new("__dotnet_date_format");
    m.arity = 2;
    m.local_count = 2;
    let s = alloc_slots(&mut m);
    let c = &mut m;

    // ── The value's parts, read once ────────────────────────────────────
    for (name, slot) in [
        ("Year", s.year),
        ("Month", s.month),
        ("Day", s.day),
        ("Hour", s.hour),
        ("Minute", s.minute),
        ("Second", s.second),
        ("Millisecond", s.millis),
    ] {
        field(c, name, line);
        set(c, slot, line);
    }

    // Day of week as 0..6. The field holds a DayOfWeek OBJECT whose
    // `Value`/`value` is the NAME (`CStr(d.DayOfWeek)` prints `Friday`), so
    // the index comes off `__index`.
    field(c, "DayOfWeek", line);
    let dow_key = c.add_constant(Value::String(Arc::from("__index")));
    c.emit_struct_field_op(Op::STRUCT_GET, 0, dow_key, line);
    set(c, s.dow, line);

    // `__offset_ms` exists on DateTimeOffset only; a DateTime is UTC-relative
    // for `z`, which .NET renders as the LOCAL offset — absent here, so 0.
    field(c, "__offset_ms", line);
    set(c, s.offset_ms, line);
    c.emit_i32_const(1, line);
    set(c, s.has_offset, line);
    get(c, s.offset_ms, line);
    c.emit_op(Op::REF_IS_NULL, line);
    c.emit_if(line);
    c.emit_f64_const(0.0, line);
    set(c, s.offset_ms, line);
    c.emit_i32_const(0, line);
    set(c, s.has_offset, line);
    c.emit_end(line);

    // 12-hour clock: `Hour % 12`, with 0 shown as 12.
    get(c, s.hour, line);
    c.emit_f64_const(12.0, line);
    vybe_compiler::primitives::math::emit_c_fmod(c, line);
    set(c, s.hour12, line);
    get(c, s.hour12, line);
    c.emit_f64_const(0.0, line);
    c.emit_op(Op::F64_EQ, line);
    c.emit_if(line);
    c.emit_f64_const(12.0, line);
    set(c, s.hour12, line);
    c.emit_end(line);

    // |offset| in minutes, for `z`/`zz`/`zzz`.
    get(c, s.offset_ms, line);
    c.emit_f64_const(60_000.0, line);
    c.emit_op(Op::F64_DIV, line);
    call(c, "ecma:math", "abs", 1, line);
    c.emit_op(Op::F64_TRUNC, line);
    set(c, s.off_abs, line);

    // ── A one-character format is a STANDARD specifier ──────────────────
    get(c, 1, line);
    set(c, s.fmt, line);
    get(c, s.fmt, line);
    str_len(c, line);
    c.emit_f64_const(1.0, line);
    c.emit_op(Op::F64_EQ, line);
    c.emit_if(line);
    for (spec, pattern) in STANDARD_PATTERNS {
        get(c, s.fmt, line);
        push_str(c, spec, line);
        str_eq(c, line);
        c.emit_if(line);
        push_str(c, pattern, line);
        set(c, s.fmt, line);
        c.emit_end(line);
    }
    c.emit_end(line);

    // ── Walk the pattern ────────────────────────────────────────────────
    push_str(c, "", line);
    set(c, s.out, line);
    c.emit_f64_const(0.0, line);
    set(c, s.i, line);
    get(c, s.fmt, line);
    str_len(c, line);
    set(c, s.len, line);

    let block_p = c.emit_block(line);
    let (loop_p, _) = c.emit_loop_s(line);

    get(c, s.i, line);
    get(c, s.len, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(c, line);
    vybe_compiler::primitives::ops::emit_dyn_not(c, line);
    c.emit_br_if(1, line);

    // ch = fmt[i]
    get(c, s.fmt, line);
    get(c, s.i, line);
    char_at(c, line);
    set(c, s.ch, line);

    // run = length of the repeat of `ch` starting at `i`
    c.emit_f64_const(1.0, line);
    set(c, s.run, line);
    let inner_block = c.emit_block(line);
    let (inner_loop, _) = c.emit_loop_s(line);
    get(c, s.i, line);
    get(c, s.run, line);
    c.emit_op(Op::F64_ADD, line);
    get(c, s.len, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(c, line);
    vybe_compiler::primitives::ops::emit_dyn_not(c, line);
    c.emit_br_if(1, line);
    get(c, s.fmt, line);
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
    c.patch_loop(inner_loop);
    c.emit_end(line);
    c.patch_block(inner_block);

    c.emit_i32_const(0, line);
    set(c, s.matched, line);

    // ── The specifier ladder ────────────────────────────────────────────
    arm(c, &s, "y", line, |c| {
        run_ge(
            c,
            &s,
            4,
            line,
            |c| {
                get(c, s.year, line);
                pad(c, 4, line);
            },
            |c| {
                get(c, s.year, line);
                c.emit_f64_const(100.0, line);
                vybe_compiler::primitives::math::emit_c_fmod(c, line);
                set(c, s.tmp, line);
                run_ge(
                    c,
                    &s,
                    2,
                    line,
                    |c| {
                        get(c, s.tmp, line);
                        pad(c, 2, line);
                    },
                    |c| {
                        get(c, s.tmp, line);
                        to_str(c, line);
                    },
                );
            },
        );
    });

    arm(c, &s, "M", line, |c| {
        // The tables are 0-based; .NET months are 1-based.
        get(c, s.month, line);
        c.emit_f64_const(1.0, line);
        c.emit_op(Op::F64_SUB, line);
        set(c, s.tmp, line);
        run_ge(
            c,
            &s,
            4,
            line,
            |c| name_at(c, MONTH_NAMES, s.tmp, NAME_WIDTH, true, line),
            |c| {
                run_ge(
                    c,
                    &s,
                    3,
                    line,
                    |c| name_at(c, MONTH_ABBR, s.tmp, 3, false, line),
                    |c| {
                        run_ge(
                            c,
                            &s,
                            2,
                            line,
                            |c| {
                                get(c, s.month, line);
                                pad(c, 2, line);
                            },
                            |c| {
                                get(c, s.month, line);
                                to_str(c, line);
                            },
                        );
                    },
                );
            },
        );
    });

    arm(c, &s, "d", line, |c| {
        run_ge(
            c,
            &s,
            4,
            line,
            |c| name_at(c, DAY_NAMES, s.dow, NAME_WIDTH, true, line),
            |c| {
                run_ge(
                    c,
                    &s,
                    3,
                    line,
                    |c| name_at(c, DAY_ABBR, s.dow, 3, false, line),
                    |c| {
                        run_ge(
                            c,
                            &s,
                            2,
                            line,
                            |c| {
                                get(c, s.day, line);
                                pad(c, 2, line);
                            },
                            |c| {
                                get(c, s.day, line);
                                to_str(c, line);
                            },
                        );
                    },
                );
            },
        );
    });

    for (spec, slot) in [
        ("H", s.hour),
        ("h", s.hour12),
        ("m", s.minute),
        ("s", s.second),
    ] {
        arm(c, &s, spec, line, |c| {
            run_ge(
                c,
                &s,
                2,
                line,
                |c| {
                    get(c, slot, line);
                    pad(c, 2, line);
                },
                |c| {
                    get(c, slot, line);
                    to_str(c, line);
                },
            );
        });
    }

    // `f`/`F` — fractional seconds, `run` digits of them. The stored
    // resolution is milliseconds, so anything past the third digit is zero;
    // .NET's `F` trims trailing zeros, `f` does not, and `"o"` asks for seven.
    for spec in ["f", "F"] {
        arm(c, &s, spec, line, |c| {
            get(c, s.millis, line);
            pad(c, 3, line);
            push_str(c, "0000", line);
            concat(c, line);
            c.emit_i32_const(0, line);
            get(c, s.run, line);
            substring(c, line);
        });
    }

    arm(c, &s, "t", line, |c| {
        get(c, s.hour, line);
        c.emit_f64_const(12.0, line);
        vybe_compiler::primitives::ops::emit_dyn_lt(c, line);
        vybe_compiler::primitives::ops::emit_dyn_to_bool(c, line);
        c.emit_if_value(line);
        push_str(c, "AM", line);
        c.emit_else(line);
        push_str(c, "PM", line);
        c.emit_end(line);
        set(c, s.tmp, line);
        run_ge(
            c,
            &s,
            2,
            line,
            |c| get(c, s.tmp, line),
            |c| {
                get(c, s.tmp, line);
                c.emit_i32_const(0, line);
                c.emit_i32_const(1, line);
                substring(c, line);
            },
        );
    });

    arm(c, &s, "z", line, |c| {
        // Sign first, then |offset| split into hours and minutes.
        get(c, s.offset_ms, line);
        c.emit_f64_const(0.0, line);
        vybe_compiler::primitives::ops::emit_dyn_lt(c, line);
        vybe_compiler::primitives::ops::emit_dyn_to_bool(c, line);
        c.emit_if_value(line);
        push_str(c, "-", line);
        c.emit_else(line);
        push_str(c, "+", line);
        c.emit_end(line);
        get(c, s.off_abs, line);
        c.emit_f64_const(60.0, line);
        c.emit_op(Op::F64_DIV, line);
        c.emit_op(Op::F64_TRUNC, line);
        set(c, s.tmp, line);
        run_ge(
            c,
            &s,
            2,
            line,
            |c| {
                get(c, s.tmp, line);
                pad(c, 2, line);
            },
            |c| {
                get(c, s.tmp, line);
                to_str(c, line);
            },
        );
        concat(c, line);
        run_ge(
            c,
            &s,
            3,
            line,
            |c| {
                push_str(c, ":", line);
                get(c, s.off_abs, line);
                c.emit_f64_const(60.0, line);
                vybe_compiler::primitives::math::emit_c_fmod(c, line);
                pad(c, 2, line);
                concat(c, line);
            },
            |c| push_str(c, "", line),
        );
        concat(c, line);
    });

    // `K` — for a `DateTimeOffset` ALWAYS the signed offset; for a `DateTime`
    // it follows `Kind` (`Z` for Utc, empty for Unspecified). The offset half
    // is what makes the round-trip pattern `"o"` reversible.
    arm(c, &s, "K", line, |c| {
        get(c, s.has_offset, line);
        c.emit_if_value(line);
        emit_offset_text(c, &s, line);
        c.emit_else(line);
        field(c, "Kind", line);
        push_str(c, "Utc", line);
        str_eq(c, line);
        c.emit_if_value(line);
        push_str(c, "Z", line);
        c.emit_else(line);
        push_str(c, "", line);
        c.emit_end(line);
        c.emit_end(line);
    });

    // `\c` — the next character, literally.
    arm(c, &s, "\\", line, |c| {
        c.emit_f64_const(2.0, line);
        set(c, s.run, line);
        get(c, s.fmt, line);
        get(c, s.i, line);
        c.emit_f64_const(1.0, line);
        c.emit_op(Op::F64_ADD, line);
        char_at(c, line);
    });

    // `%c` — "the SINGLE custom specifier c", which is what a bare `c` already
    // means inside a multi-character pattern. Consume the `%` and let the next
    // iteration read `c` with its own run of 1.
    arm(c, &s, "%", line, |c| {
        c.emit_f64_const(1.0, line);
        set(c, s.run, line);
        push_str(c, "", line);
    });

    // `'…'` and `"…"` — a literal section, verbatim.
    for quote in ["'", "\""] {
        arm(c, &s, quote, line, |c| {
            get(c, s.fmt, line);
            push_str(c, quote, line);
            get(c, s.i, line);
            c.emit_f64_const(1.0, line);
            c.emit_op(Op::F64_ADD, line);
            call(c, "ecma:string", "indexOf", 3, line);
            set(c, s.tmp, line);
            get(c, s.tmp, line);
            c.emit_f64_const(0.0, line);
            vybe_compiler::primitives::ops::emit_dyn_lt(c, line);
            vybe_compiler::primitives::ops::emit_dyn_to_bool(c, line);
            c.emit_if(line);
            // Unterminated: the rest of the pattern is the literal.
            get(c, s.len, line);
            set(c, s.tmp, line);
            c.emit_end(line);
            get(c, s.tmp, line);
            get(c, s.i, line);
            c.emit_op(Op::F64_SUB, line);
            c.emit_f64_const(1.0, line);
            c.emit_op(Op::F64_ADD, line);
            set(c, s.run, line);
            get(c, s.fmt, line);
            get(c, s.i, line);
            c.emit_f64_const(1.0, line);
            c.emit_op(Op::F64_ADD, line);
            get(c, s.tmp, line);
            substring(c, line);
        });
    }

    // Anything else is a literal separator — `-`, `/`, `:`, a space.
    get(c, s.matched, line);
    c.emit_op(Op::I32_EQZ, line);
    c.emit_if(line);
    get(c, s.ch, line);
    set(c, s.piece, line);
    c.emit_end(line);

    get(c, s.out, line);
    get(c, s.piece, line);
    concat(c, line);
    set(c, s.out, line);

    get(c, s.i, line);
    get(c, s.run, line);
    c.emit_op(Op::F64_ADD, line);
    set(c, s.i, line);

    c.emit_br(0, line);
    c.emit_end(line);
    c.patch_loop(loop_p);
    c.emit_end(line);
    c.patch_block(block_p);

    get(c, s.out, line);
    c.emit_op(Op::RETURN, line);
    m
}

/// The interpreter's chunk index, created once per module.
fn format_chunk(chunks: &mut Vec<Chunk>, line: u32) -> usize {
    if let Some(idx) = chunks.iter().position(|c| c.name == "__dotnet_date_format") {
        return idx;
    }
    chunks.push(build_chunk(line));
    chunks.len() - 1
}

/// `value.ToString(format)`. Stack: `[value, format] → [string]`.
pub fn emit_date_format(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let idx = format_chunk(chunks, line);
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::REF_FUNC, idx as u16, line);
    chunk.emit(0, line);
    let value_slot = chunk.alloc_scratch(3);
    let fmt_slot = value_slot + 1;
    let fn_slot = value_slot + 2;
    chunk.emit_op_u16(Op::LOCAL_SET, fn_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, fmt_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, value_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, fn_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, fmt_slot, line);
    chunk.emit_op_u8_u8(Op::CALL_REF, 2, 1, line);
}
