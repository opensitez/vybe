//! `System.Char`'s static surface, shared by every .NET language.
//!
//! Eight classifiers were already registered (`IsDigit`, `IsLetter`,
//! `IsLetterOrDigit`, `IsUpper`, `IsLower`, `IsWhiteSpace`, `ToUpper`,
//! `ToLower`) as arity-1 `MethodDef`s over the shared `primitives::strings`
//! predicates. This file adds the rest of the surface and — the part that
//! cannot be fixed by adding more rows — the OVERLOADS.
//!
//! ⛔ Static method lookup keys on the NAME ALONE. `emitter/mod.rs`'s instance
//! path matches `name` *and* `arity`, but the static path (`mod.rs:574`) never
//! looks at arity, so registering `IsLetter/1` and `IsLetter/2` does not give
//! two overloads — the first registration wins and the second is dead. Every
//! overloaded name here is therefore ONE registration whose body branches on
//! `argc`, the same shape `lazy_adapter::emit_lazy_new` uses.
//!
//! A char IS a one-character string at runtime (there is no `ecma:char` host
//! module and never was — ECMAScript has no char type). So the `(char)` and
//! `(string, index)` overloads differ only in how the character is REACHED:
//! `narrow_to_char` reduces both to the same one-character string and hands off
//! to the shared Unicode-aware classifier, rather than this file restating
//! Unicode tables it would get wrong.

use vybe_compiler::primitives::ops::emit_i32_to_bool;
use vybe_compiler::primitives::strings;
use vybe_runtime::opcode::Op;
use vybe_runtime::Chunk;

/// `.NET` classifies ASCII punctuation and symbols by Unicode CATEGORY, and the
/// split is not derivable from "not alphanumeric, not space": `.` is `Po` but
/// `+` is `Sm`. Measured against `tools/vbrun`, `Char.IsPunctuation("+"c)` is
/// False and `Char.IsSymbol("+"c)` is True. For ASCII the two sets are small
/// and closed, so they are spelled out rather than guessed at.
const ASCII_PUNCTUATION: &[i32] = &[
    0x21, 0x22, 0x23, 0x25, 0x26, 0x27, 0x28, 0x29, 0x2A, 0x2C, 0x2D, 0x2E, 0x2F, 0x3A, 0x3B, 0x3F,
    0x40, 0x5B, 0x5C, 0x5D, 0x5F, 0x7B, 0x7D,
];

/// `$ + < = > ^ \` | ~` — `Sc`, `Sm`, `Sk`, `So`.
const ASCII_SYMBOL: &[i32] = &[0x24, 0x2B, 0x3C, 0x3D, 0x3E, 0x5E, 0x60, 0x7C, 0x7E];

/// `Zs`/`Zl`/`Zp`. NOT the same as `IsWhiteSpace`: a tab is white space but is
/// not a separator, which is exactly what the .NET pair asserts.
const SEPARATORS: &[(i32, i32)] = &[
    (0x20, 0x20),
    (0xA0, 0xA0),
    (0x1680, 0x1680),
    (0x2000, 0x200A),
    (0x2028, 0x2029),
    (0x202F, 0x202F),
    (0x205F, 0x205F),
    (0x3000, 0x3000),
];

fn get(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn set(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}

/// `charCodeAt`. Stack: `[string, index]` → `[i32]`.
fn char_code_at(chunk: &mut Chunk, line: u32) {
    let f = chunk.add_import("wasm:js-string", "charCodeAt");
    chunk.emit_call(f, 2, line);
}

/// Reduce the arguments of a classifier call to the UTF-16 CODE UNIT.
/// `argc == 1`: `[char]`. `argc >= 2`: `[string, index]`.
///
/// The single-argument path goes through `strings::emit_char_code` rather than
/// a raw `charCodeAt`: the primitive tests `wasm:js-string.test` first and
/// coerces a non-string through `ecma:number`. That guard is not decoration —
/// `Char.MaxValue` is a NUMBER in this tree (the constant table is f64-only),
/// so a raw `charCodeAt` traps on exactly the values a limit test passes in.
fn narrow_to_code(chunk: &mut Chunk, argc: u8, line: u32) {
    if argc < 2 {
        strings::emit_char_code(chunk, line);
        return;
    }
    // No shared equivalent for the INDEXED read: `emit_code_point_at` is
    // `codePointAt`, which pairs surrogates — right for a code point, wrong for
    // `Char.IsSurrogatePair(s, i)`, whose whole job is to see the halves.
    char_code_at(chunk, line);
}

/// Reduce the same arguments to a ONE-CHARACTER STRING, so the call can hand
/// off to a shared `primitives::strings` predicate unchanged.
fn narrow_to_char(chunk: &mut Chunk, argc: u8, line: u32) {
    if argc >= 2 {
        char_code_at(chunk, line);
        strings::emit_from_char_code(chunk, line);
    }
}

/// `code >= lo && code <= hi`, reading `slot`. Leaves an i32 0/1.
fn in_range(chunk: &mut Chunk, slot: u16, lo: i32, hi: i32, line: u32) {
    if lo == hi {
        get(chunk, slot, line);
        chunk.emit_i32_const(lo, line);
        chunk.emit_op(Op::I32_EQ, line);
        return;
    }
    get(chunk, slot, line);
    chunk.emit_i32_const(lo, line);
    chunk.emit_op(Op::I32_GE_S, line);
    get(chunk, slot, line);
    chunk.emit_i32_const(hi, line);
    chunk.emit_op(Op::I32_LE_S, line);
    chunk.emit_op(Op::I32_AND, line);
}

/// OR of a range set, reading `slot`. Leaves an i32 0/1.
fn in_ranges(chunk: &mut Chunk, slot: u16, ranges: &[(i32, i32)], line: u32) {
    for (i, (lo, hi)) in ranges.iter().enumerate() {
        in_range(chunk, slot, *lo, *hi, line);
        if i > 0 {
            chunk.emit_op(Op::I32_OR, line);
        }
    }
}

fn in_set(chunk: &mut Chunk, slot: u16, codes: &[i32], line: u32) {
    for (i, code) in codes.iter().enumerate() {
        in_range(chunk, slot, *code, *code, line);
        if i > 0 {
            chunk.emit_op(Op::I32_OR, line);
        }
    }
}

/// The shape every code-point predicate shares: narrow the arguments to a code
/// unit, park it, test it, lift the result to a real Boolean.
///
/// The lift matters: .NET returns `Boolean` and VB renders that `True`/`False`,
/// so a bare i32 would print `1`.
fn code_predicate(
    chunks: &mut [Chunk],
    current: usize,
    argc: u8,
    line: u32,
    test: impl FnOnce(&mut Chunk, u16, u32),
) {
    let chunk = &mut chunks[current];
    narrow_to_code(chunk, argc, line);
    let slot = chunk.alloc_scratch(1);
    set(chunk, slot, line);
    test(chunk, slot, line);
    emit_i32_to_bool(chunk, line);
}

// ── Classifiers that delegate to the shared Unicode predicates ──────────────
//
// These exist ONLY to add the `(string, index)` overload. The single-argument
// path is byte-for-byte what the previous arity-1 registration emitted.

macro_rules! delegating {
    ($(#[$doc:meta])* $name:ident => $shared:path) => {
        $(#[$doc])*
        pub fn $name(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
            narrow_to_char(&mut chunks[current], argc, line);
            $shared(chunks, current, line);
        }
    };
}

delegating!(
    /// `Char.IsDigit(c)` / `Char.IsDigit(s, i)`.
    emit_is_digit => strings::emit_is_digit
);
delegating!(
    /// `Char.IsLetter(c)` / `Char.IsLetter(s, i)`.
    emit_is_letter => strings::emit_is_alpha
);
delegating!(
    /// `Char.IsLetterOrDigit(c)` / `Char.IsLetterOrDigit(s, i)`.
    emit_is_letter_or_digit => strings::emit_is_alnum
);
delegating!(
    /// `Char.IsUpper(c)` / `Char.IsUpper(s, i)`.
    emit_is_upper => strings::emit_is_upper
);
delegating!(
    /// `Char.IsLower(c)` / `Char.IsLower(s, i)`.
    emit_is_lower => strings::emit_is_lower
);
delegating!(
    /// `Char.IsWhiteSpace(c)` / `Char.IsWhiteSpace(s, i)`.
    emit_is_white_space => strings::emit_is_space
);

// ── Case conversion ─────────────────────────────────────────────────────────

/// `Char.ToUpper(c)` and `Char.ToUpper(c, culture)`.
///
/// The culture argument is DROPPED, not honoured: the only culture the tests
/// name is `InvariantCulture`, and `ecma:string`'s `toUpperCase` is already the
/// invariant mapping. A culture-sensitive mapping (Turkish dotless ı) would be
/// a different implementation, and silently pretending to honour the argument
/// while ignoring it is the failure mode worth being loud about — so this is
/// recorded here rather than hidden.
pub fn emit_to_upper(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    if argc >= 2 {
        chunk.emit_op(Op::DROP, line);
    }
    strings::emit_to_upper(chunk, line);
}

/// `Char.ToLower(c)` and `Char.ToLower(c, culture)`. See [`emit_to_upper`].
pub fn emit_to_lower(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    if argc >= 2 {
        chunk.emit_op(Op::DROP, line);
    }
    strings::emit_to_lower(chunk, line);
}

// ── Code-point predicates ───────────────────────────────────────────────────

/// `Char.IsAscii(c)` — a single code unit below U+0080.
pub fn emit_is_ascii(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    code_predicate(chunks, current, argc, line, |c, slot, line| {
        in_range(c, slot, 0, 0x7F, line)
    });
}

/// `Char.IsAsciiDigit(c)`.
pub fn emit_is_ascii_digit(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    code_predicate(chunks, current, argc, line, |c, slot, line| {
        in_range(c, slot, 0x30, 0x39, line)
    });
}

/// `Char.IsAsciiLetter(c)`.
pub fn emit_is_ascii_letter(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    code_predicate(chunks, current, argc, line, |c, slot, line| {
        in_ranges(c, slot, &[(0x41, 0x5A), (0x61, 0x7A)], line)
    });
}

/// `Char.IsAsciiLetterOrDigit(c)`.
pub fn emit_is_ascii_letter_or_digit(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    code_predicate(chunks, current, argc, line, |c, slot, line| {
        in_ranges(c, slot, &[(0x30, 0x39), (0x41, 0x5A), (0x61, 0x7A)], line)
    });
}

/// `Char.IsAsciiHexDigit(c)` — `0-9`, `A-F`, `a-f`.
pub fn emit_is_ascii_hex_digit(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    code_predicate(chunks, current, argc, line, |c, slot, line| {
        in_ranges(c, slot, &[(0x30, 0x39), (0x41, 0x46), (0x61, 0x66)], line)
    });
}

/// `Char.IsControl(c)` — C0 and C1.
pub fn emit_is_control(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    code_predicate(chunks, current, argc, line, |c, slot, line| {
        in_ranges(c, slot, &[(0x00, 0x1F), (0x7F, 0x9F)], line)
    });
}

/// `Char.IsSeparator(c)` — `Zs`/`Zl`/`Zp`.
pub fn emit_is_separator(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    code_predicate(chunks, current, argc, line, |c, slot, line| {
        in_ranges(c, slot, SEPARATORS, line)
    });
}

/// `Char.IsPunctuation(c)` — ASCII only; see [`ASCII_PUNCTUATION`].
pub fn emit_is_punctuation(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    code_predicate(chunks, current, argc, line, |c, slot, line| {
        in_set(c, slot, ASCII_PUNCTUATION, line)
    });
}

/// `Char.IsSymbol(c)` — ASCII only; see [`ASCII_SYMBOL`].
pub fn emit_is_symbol(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    code_predicate(chunks, current, argc, line, |c, slot, line| {
        in_set(c, slot, ASCII_SYMBOL, line)
    });
}

/// `Char.IsHighSurrogate(c)` — U+D800..U+DBFF.
pub fn emit_is_high_surrogate(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    code_predicate(chunks, current, argc, line, |c, slot, line| {
        in_range(c, slot, 0xD800, 0xDBFF, line)
    });
}

/// `Char.IsLowSurrogate(c)` — U+DC00..U+DFFF.
pub fn emit_is_low_surrogate(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    code_predicate(chunks, current, argc, line, |c, slot, line| {
        in_range(c, slot, 0xDC00, 0xDFFF, line)
    });
}

/// `Char.IsSurrogate(c)` — either half.
pub fn emit_is_surrogate(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    code_predicate(chunks, current, argc, line, |c, slot, line| {
        in_range(c, slot, 0xD800, 0xDFFF, line)
    });
}

// ── The two-argument conversions ────────────────────────────────────────────

/// `Char.IsSurrogatePair(high, low)` and `Char.IsSurrogatePair(s, index)`.
///
/// ⛔ The one method in this file that argc CANNOT discriminate: both overloads
/// take two arguments and differ only in the TYPE of the second. The runtime
/// test is `wasm:js-string.test` on argument 2 — a string means the
/// `(char, char)` form, a number means `(string, index)` and the low half is at
/// `index + 1`.
pub fn emit_is_surrogate_pair(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let second = chunk.alloc_scratch(1);
    let first = chunk.alloc_scratch(1);
    let high = chunk.alloc_scratch(1);
    let low = chunk.alloc_scratch(1);
    set(chunk, second, line);
    set(chunk, first, line);

    get(chunk, second, line);
    let is_str = chunk.add_import("wasm:js-string", "test");
    chunk.emit_call(is_str, 1, line);
    chunk.emit_if_value(line);

    // `(char, char)` — each argument is its own one-character string.
    get(chunk, first, line);
    chunk.emit_i32_const(0, line);
    char_code_at(chunk, line);
    set(chunk, high, line);
    get(chunk, second, line);
    chunk.emit_i32_const(0, line);
    char_code_at(chunk, line);
    set(chunk, low, line);
    chunk.emit_i32_const(0, line);

    chunk.emit_else(line);

    // `(string, index)` — the pair is at `index` and `index + 1`.
    get(chunk, first, line);
    get(chunk, second, line);
    char_code_at(chunk, line);
    set(chunk, high, line);
    get(chunk, first, line);
    get(chunk, second, line);
    chunk.emit_i32_const(1, line);
    chunk.emit_op(Op::I32_ADD, line);
    char_code_at(chunk, line);
    set(chunk, low, line);
    chunk.emit_i32_const(0, line);

    chunk.emit_end(line);
    chunk.emit_op(Op::DROP, line);

    in_range(chunk, high, 0xD800, 0xDBFF, line);
    in_range(chunk, low, 0xDC00, 0xDFFF, line);
    chunk.emit_op(Op::I32_AND, line);
    emit_i32_to_bool(chunk, line);
}

/// `Char.ConvertToUtf32(high, low)` — the astral code point a surrogate pair
/// names. `(high - 0xD800) * 0x400 + (low - 0xDC00) + 0x10000`.
pub fn emit_convert_to_utf32(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let low = chunk.alloc_scratch(1);
    let high = chunk.alloc_scratch(1);
    set(chunk, low, line);
    set(chunk, high, line);

    get(chunk, high, line);
    chunk.emit_i32_const(0, line);
    char_code_at(chunk, line);
    chunk.emit_i32_const(0xD800, line);
    chunk.emit_op(Op::I32_SUB, line);
    chunk.emit_i32_const(0x400, line);
    chunk.emit_op(Op::I32_MUL, line);

    get(chunk, low, line);
    chunk.emit_i32_const(0, line);
    char_code_at(chunk, line);
    chunk.emit_i32_const(0xDC00, line);
    chunk.emit_op(Op::I32_SUB, line);
    chunk.emit_op(Op::I32_ADD, line);

    chunk.emit_i32_const(0x10000, line);
    chunk.emit_op(Op::I32_ADD, line);
    chunk.emit_op(Op::F64_FROM_I32, line);
}

/// `Char.ConvertFromUtf32(cp)` — the (possibly surrogate) STRING for a code
/// point. This is `fromCodePoint`, not `fromCharCode`: the whole point is that
/// an astral code point becomes the two-unit pair.
pub fn emit_convert_from_utf32(chunks: &mut [Chunk], current: usize, line: u32) {
    strings::emit_from_code_point(&mut chunks[current], line);
}

/// `Char.GetNumericValue(c)` — the numeric value of a digit, `-1` otherwise.
///
/// Scoped to ASCII digits. The full `Nd`/`No` table (Roman numerals, vulgar
/// fractions, every script's digits) is a Unicode database, not a rule, and
/// nothing measured needs it — `-1` is the correct answer for everything else.
pub fn emit_get_numeric_value(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    narrow_to_code(chunk, argc, line);
    let code = chunk.alloc_scratch(1);
    set(chunk, code, line);

    in_range(chunk, code, 0x30, 0x39, line);
    chunk.emit_if_value(line);
    get(chunk, code, line);
    chunk.emit_i32_const(0x30, line);
    chunk.emit_op(Op::I32_SUB, line);
    chunk.emit_op(Op::F64_FROM_I32, line);
    chunk.emit_else(line);
    chunk.emit_f64_const(-1.0, line);
    chunk.emit_end(line);
}

/// `Char.GetUnicodeCategory(c)` — the category NAME.
///
/// .NET returns a `UnicodeCategory` enum whose `ToString()` is this name, and a
/// string answers `.ToString()` as itself, so the name IS the value. This is a
/// LADDER over the categories the shared classifiers and the ASCII sets above
/// can actually distinguish — deliberately not a Unicode database. Anything it
/// cannot place answers `OtherSymbol`, which is wrong-but-honest rather than a
/// fabricated table that would be wrong in more places.
pub fn emit_get_unicode_category(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    {
        let chunk = &mut chunks[current];
        narrow_to_code(chunk, argc, line);
    }
    let code = chunks[current].alloc_scratch(1);
    {
        let chunk = &mut chunks[current];
        set(chunk, code, line);
    }

    // Ordered most-specific first; every arm pushes exactly one string.
    let ladder: &[(&[(i32, i32)], &str)] = &[
        (&[(0x30, 0x39)], "DecimalDigitNumber"),
        (&[(0x41, 0x5A)], "UppercaseLetter"),
        (&[(0x61, 0x7A)], "LowercaseLetter"),
        (&[(0x00, 0x1F), (0x7F, 0x9F)], "Control"),
        (SEPARATORS, "SpaceSeparator"),
    ];

    let mut opened = 0;
    for (ranges, name) in ladder {
        let chunk = &mut chunks[current];
        in_ranges(chunk, code, ranges, line);
        chunk.emit_if_value(line);
        chunk.emit_string_const(name, line);
        chunk.emit_else(line);
        opened += 1;
    }

    {
        let chunk = &mut chunks[current];
        in_set(chunk, code, ASCII_PUNCTUATION, line);
        chunk.emit_if_value(line);
        chunk.emit_string_const("OtherPunctuation", line);
        chunk.emit_else(line);
        chunk.emit_string_const("OtherSymbol", line);
        chunk.emit_end(line);
    }

    for _ in 0..opened {
        chunks[current].emit_end(line);
    }
}

// ────────────────────────────────────────────────────────────────────────────
// `System.Text.Rune` — one Unicode SCALAR value.
// ────────────────────────────────────────────────────────────────────────────

/// `new Rune(char)` / `new Rune(int)`.
///
/// A Rune is a value struct whose whole state is the scalar value, and every
/// other member is a pure function of it — so they are COMPUTED ONCE and stored
/// as fields rather than reached through property machinery. `System.Index` is
/// registered the same way.
///
/// The lengths are the encoded sizes, read off the SDK: U+0041 is 1 UTF-8 byte
/// and 1 UTF-16 unit; U+1F600 is 4 and 2.
pub fn emit_rune_new(chunks: &mut [Chunk], current: usize, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    // `emit_char_code` takes either spelling: it tests for a string first and
    // coerces a bare number through `ecma:number`, which is what lets one
    // constructor serve `Rune('A')` and `Rune(0x1F600)`.
    strings::emit_char_code(&mut chunks[current], line);
    set(&mut chunks[current], value, line);

    get(&mut chunks[current], value, line);
    rune_is_ascii(&mut chunks[current], value, line);
    rune_is_bmp(&mut chunks[current], value, line);
    rune_utf8_length(&mut chunks[current], value, line);
    rune_utf16_length(&mut chunks[current], value, line);
    crate::emitter::dispatch::emit_value_type_new(
        &mut chunks[current],
        "Rune",
        &[
            "Value",
            "IsAscii",
            "IsBmp",
            "Utf8SequenceLength",
            "Utf16SequenceLength",
        ],
        line,
    );
}

fn rune_is_ascii(chunk: &mut Chunk, value: u16, line: u32) {
    get(chunk, value, line);
    chunk.emit_i32_const(0x80, line);
    chunk.emit_op(Op::I32_LT_S, line);
    emit_i32_to_bool(chunk, line);
}

fn rune_is_bmp(chunk: &mut Chunk, value: u16, line: u32) {
    get(chunk, value, line);
    chunk.emit_i32_const(0x10000, line);
    chunk.emit_op(Op::I32_LT_S, line);
    emit_i32_to_bool(chunk, line);
}

/// 1 below U+0080, 2 below U+0800, 3 below U+10000, else 4.
fn rune_utf8_length(chunk: &mut Chunk, value: u16, line: u32) {
    for (limit, len) in [(0x80, 1), (0x800, 2), (0x10000, 3)] {
        get(chunk, value, line);
        chunk.emit_i32_const(limit, line);
        chunk.emit_op(Op::I32_LT_S, line);
        chunk.emit_if_value(line);
        chunk.emit_f64_const(len as f64, line);
        chunk.emit_else(line);
    }
    chunk.emit_f64_const(4.0, line);
    for _ in 0..3 {
        chunk.emit_end(line);
    }
}

/// A scalar outside the BMP needs a surrogate PAIR.
fn rune_utf16_length(chunk: &mut Chunk, value: u16, line: u32) {
    get(chunk, value, line);
    chunk.emit_i32_const(0x10000, line);
    chunk.emit_op(Op::I32_LT_S, line);
    chunk.emit_if_value(line);
    chunk.emit_f64_const(1.0, line);
    chunk.emit_else(line);
    chunk.emit_f64_const(2.0, line);
    chunk.emit_end(line);
}

/// Read `Value` off a Rune receiver, leaving the scalar on the stack.
///
/// The static classifiers are declared on `Rune` and take a Rune, but the code
/// predicates above already answer for a scalar — `emit_char_code` coerces a
/// bare number — so each static is the unwrap plus the classifier it shares
/// with `System.Char`, not a second copy of the tables.
fn rune_value(chunks: &mut [Chunk], current: usize, line: u32) {
    vybe_compiler::primitives::class_slots::emit_class_get(
        &mut chunks[current],
        vybe_compiler::primitives::class_slots::ObjSource::Stack,
        &super::object_fields::field_slot("Value"),
        vybe_compiler::primitives::class_slots::Dest::Stack,
        line,
    );
}

/// `Rune.ToString()` — the character the scalar denotes, surrogate pair and all.
pub fn emit_rune_to_string(chunks: &mut [Chunk], current: usize, line: u32) {
    rune_value(chunks, current, line);
    strings::emit_from_code_point(&mut chunks[current], line);
}

/// `Rune.GetHashCode()` — the scalar itself, which is what the SDK answers.
pub fn emit_rune_hash(chunks: &mut [Chunk], current: usize, line: u32) {
    rune_value(chunks, current, line);
}

/// `a.CompareTo(b)` — ordinal on the scalar: negative, zero or positive.
pub fn emit_rune_compare_to(chunks: &mut [Chunk], current: usize, line: u32) {
    let other = chunks[current].alloc_scratch(1);
    let mine = chunks[current].alloc_scratch(1);
    rune_value(chunks, current, line);
    set(&mut chunks[current], other, line);
    rune_value(chunks, current, line);
    set(&mut chunks[current], mine, line);
    get(&mut chunks[current], mine, line);
    get(&mut chunks[current], other, line);
    chunks[current].emit_op(Op::F64_SUB, line);
    // .NET answers −1/0/1 rather than the raw difference.
    let sign = chunks[current].add_import("ecma:math", "sign");
    chunks[current].emit_call(sign, 1, line);
}

/// `a.Equals(b)` — two Runes are equal exactly when their scalars are.
pub fn emit_rune_equals(chunks: &mut [Chunk], current: usize, line: u32) {
    let other = chunks[current].alloc_scratch(1);
    rune_value(chunks, current, line);
    set(&mut chunks[current], other, line);
    rune_value(chunks, current, line);
    get(&mut chunks[current], other, line);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    emit_i32_to_bool(&mut chunks[current], line);
}

/// A `Rune` static classifier: unwrap the scalar, then run the shared code
/// predicate named by `kind`.
pub fn emit_rune_predicate(chunks: &mut [Chunk], current: usize, kind: &str, line: u32) {
    rune_value(chunks, current, line);
    strings::emit_from_code_point(&mut chunks[current], line);
    match kind {
        "digit" => strings::emit_is_digit(chunks, current, line),
        "letter" => strings::emit_is_alpha(chunks, current, line),
        "letterordigit" => strings::emit_is_alnum(chunks, current, line),
        "upper" => strings::emit_is_upper(chunks, current, line),
        "lower" => strings::emit_is_lower(chunks, current, line),
        "whitespace" => strings::emit_is_space(chunks, current, line),
        "control" => emit_is_control(chunks, current, 1, line),
        _ => emit_is_punctuation(chunks, current, 1, line),
    }
}

/// `Rune.ToUpperInvariant(r)` / `Rune.ToLowerInvariant(r)` — a Rune in, a Rune
/// out, so the case fold runs on the character and the result is re-wrapped.
pub fn emit_rune_case(chunks: &mut [Chunk], current: usize, upper: bool, line: u32) {
    rune_value(chunks, current, line);
    strings::emit_from_code_point(&mut chunks[current], line);
    if upper {
        strings::emit_to_upper(&mut chunks[current], line);
    } else {
        strings::emit_to_lower(&mut chunks[current], line);
    }
    emit_rune_new(chunks, current, line);
}

/// `Rune.ReplacementChar` — U+FFFD, the scalar a decoder substitutes.
pub fn emit_rune_replacement_char(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_f64_const(0xFFFD as f64, line);
    emit_rune_new(chunks, current, line);
}

/// `Rune.GetNumericValue(r)` — the digit's value, or −1 when it has none.
pub fn emit_rune_numeric_value(chunks: &mut [Chunk], current: usize, line: u32) {
    let code = chunks[current].alloc_scratch(1);
    rune_value(chunks, current, line);
    set(&mut chunks[current], code, line);
    in_range(&mut chunks[current], code, 0x30, 0x39, line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], code, line);
    chunks[current].emit_f64_const(0x30 as f64, line);
    chunks[current].emit_op(Op::F64_SUB, line);
    chunks[current].emit_else(line);
    chunks[current].emit_f64_const(-1.0, line);
    chunks[current].emit_end(line);
}
