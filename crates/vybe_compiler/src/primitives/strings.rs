//! String interpolation helpers — shared bytecode patterns for string building.
//!
//! All languages (Python f-strings, Dart $interpolation, JS template literals,
//! C# $strings, VB string concat) emit the same pattern: compile parts,
//! toString each expression, concatenate.

use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

/// Emit toString conversion on TOS via host call.
/// Stack before: [value]  Stack after: [string]
/// Emit toString conversion. Adds import to the given chunk.
/// Use `emit_to_string_with_import` if your compiler requires imports in chunk 0.
pub fn emit_to_string(chunk: &mut Chunk, line: u32) {
    let to_str = chunk.add_import("ecma:string", "String");
    chunk.emit_call(to_str, 1, line);
}

/// Emit toString using a pre-resolved import index.
pub fn emit_to_string_with_import(chunk: &mut Chunk, import_idx: u16, line: u32) {
    chunk.emit_call(import_idx, 1, line);
}

/// Emit concatenation of N string parts already on the stack.
/// If N == 0, pushes empty string. If N == 1, no-op.
/// Stack before: [part1, part2, ..., partN]  Stack after: [concatenated_string]
pub fn emit_concat(chunk: &mut Chunk, part_count: usize, line: u32) {
    if part_count == 0 {
        chunk.emit_string_const("", line);
    } else if part_count > 1 {
        let concat_idx = chunk.add_import("wasm:js-string", "concat");
        let base = chunk.local_count;
        chunk.local_count = chunk
            .local_count
            .checked_add(part_count as u16)
            .expect("emit_concat: local slot overflow");
        for i in (0..part_count).rev() {
            chunk.emit_op_u16(Op::LOCAL_SET, base + i as u16, line);
        }
        chunk.emit_op_u16(Op::LOCAL_GET, base, line);
        for i in 1..part_count {
            chunk.emit_op_u16(Op::LOCAL_GET, base + i as u16, line);
            chunk.emit_call(concat_idx, 2, line);
        }
    }
    // part_count == 1: string is already on stack, nothing to do
}

/// Emit a complete string interpolation: convert expression part to string.
/// Call this after compiling each expression part (between literal parts).
/// Stack before: [value]  Stack after: [string]
pub fn emit_interpolation_part(chunk: &mut Chunk, line: u32) {
    emit_to_string(chunk, line);
}

/// Emit a string literal part.
/// Stack: [] → [string]
pub fn emit_literal_part(chunk: &mut Chunk, text: &str, line: u32) {
    chunk.emit_string_const(text, line);
}

// ── String operations ──────────────────────────────────────────────────
// Single-opcode wrappers for consistency across all compilers.

/// String length. Stack: [string] → [i32]
pub fn emit_length(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("wasm:js-string", "length");
    chunk.emit_call(idx, 1, line);
}

/// Substring. Stack: [string, start, length] → [string]
pub fn emit_substring(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("wasm:js-string", "substring");
    chunk.emit_call(idx, 3, line);
}

/// Index of substring. Stack: [haystack, needle] → [i32]
pub fn emit_index_of(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("ecma:string", "indexOf");
    chunk.emit_call(idx, 2, line);
}

/// Last index of substring. Stack: [haystack, needle] → [i32]
pub fn emit_last_index_of(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("ecma:string", "lastIndexOf");
    chunk.emit_call(idx, 2, line);
}

/// Replace. Stack: [string, search, replace] → [string]
pub fn emit_replace(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("ecma:string", "replaceAll");
    chunk.emit_call(idx, 3, line);
}

/// Split. Stack: [string, delimiter] → [array]
pub fn emit_split(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("ecma:string", "split");
    chunk.emit_call(idx, 2, line);
}

/// To lowercase. Stack: [string] → [string]
pub fn emit_to_lower(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("ecma:string", "toLowerCase");
    chunk.emit_call(idx, 1, line);
}

/// To uppercase. Stack: [string] → [string]
pub fn emit_to_upper(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("ecma:string", "toUpperCase");
    chunk.emit_call(idx, 1, line);
}

/// Trim whitespace. Stack: [string] → [string]
pub fn emit_trim(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("ecma:string", "trim");
    chunk.emit_call(idx, 1, line);
}

/// Trim start. Stack: [string] → [string]
pub fn emit_trim_start(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("ecma:string", "trimStart");
    chunk.emit_call(idx, 1, line);
}

/// Trim end. Stack: [string] → [string]
pub fn emit_trim_end(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("ecma:string", "trimEnd");
    chunk.emit_call(idx, 1, line);
}

/// Repeat string. Stack: [string, count] → [string]
pub fn emit_repeat(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("ecma:string", "repeat");
    chunk.emit_call(idx, 2, line);
}

/// Pairwise concatenation. Stack: [a, b] → [ab]
pub fn emit_str_concat(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("wasm:js-string", "concat");
    chunk.emit_call(idx, 2, line);
}

/// Concat with ToString coercion of both operands (VB `&`, C# string
/// concat, …): `wasm:js-string.concat` is spec-strict and traps on
/// non-string args, so operands that may be numbers/booleans must go
/// through `ecma:string.String` first.
/// Stack: [l, r] → [String(l) + String(r)]
pub fn emit_str_concat_coercing(chunk: &mut Chunk, line: u32) {
    let to_str = chunk.add_import("ecma:string", "String");
    let concat = chunk.add_import("wasm:js-string", "concat");
    let r = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, r, line); // [l]
    chunk.emit_call(to_str, 1, line); // [String(l)]
    chunk.emit_op_u16(Op::LOCAL_GET, r, line); // [String(l), r]
    chunk.emit_call(to_str, 1, line); // [String(l), String(r)]
    chunk.emit_call(concat, 2, line);
}

/// Reverse string. Stack: [string] → [reversed]
/// Composed: split("") → reverse() → join("")
pub fn emit_str_reverse(chunk: &mut Chunk, line: u32) {
    chunk.emit_string_const("", line);
    let split = chunk.add_import("ecma:string", "split");
    chunk.emit_call(split, 2, line);
    let reverse = chunk.add_import("ecma:array", "reverse");
    chunk.emit_call(reverse, 1, line);
    chunk.emit_string_const("", line);
    let join = chunk.add_import("ecma:array", "join");
    chunk.emit_call(join, 2, line);
}

// ── Scalar (Unicode code point) unit ───────────────────────────────────
//
// `unifiedstringplan.md` Axis 1: `length`, `[i]`, `substring` and `indexOf`
// count in a unit, and which unit is a per-language property. The helpers
// ABOVE all count in **UTF-16 code units** — `wasm:js-string.length` and
// `ecma:string.length` are both `s.encode_utf16().count()`, verified at their
// registration sites, so the whole `str_*` surface is one consistent unit.
//
// These are the **`scalar`** counterparts, for the languages and library
// surfaces that count Unicode code points: Python `str`, PHP's `mb_*` family,
// Lua's `utf8` library. They differ from the UTF-16 helpers only outside the
// BMP, where one code point is two UTF-16 units — which is exactly where the
// UTF-16 helpers cut an astral character in half and yield a replacement char.
//
// Everything here composes primitives that already exist. `ecma:array.from`
// walks a string with the **string iterator** (ECMA-262 §22.1.3.34), which
// yields code points, so `Array.from("😀").length` is 1 — that is the
// `[...s]` idiom, and it is why no host function was needed.
//
// NOT provided, deliberately: the `byte` unit. PHP `strlen`/`substr`, Lua `#`
// and Go `len` count UTF-8 bytes, and a byte substring can cut a code point in
// half — which `Value::String` (`Arc<str>`) cannot represent. That half is
// blocked on `Literal::Bytes` (plan §3c), not merely unwritten.

/// The code points of a string, as an array of one-code-point strings.
/// Stack: `[string]` → `[array]`. This is `[...s]`, and PHP `mb_str_split`
/// with the default chunk length.
pub fn emit_scalar_chars(chunk: &mut Chunk, line: u32) {
    let from = chunk.add_import("ecma:array", "from");
    chunk.emit_call(from, 1, line);
}

/// Code-point length. Stack: `[string]` → `[i32]`.
/// `mb_strlen("a😀b")` is 3 where the UTF-16 count is 4.
pub fn emit_scalar_length(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_scalar_chars(&mut chunks[current], line);
    crate::primitives::collections::emit_len(chunks, current, line);
}

/// Code-point substring. Stack: `[string, start, end]` → `[string]`, `end`
/// EXCLUSIVE. Bounds wrap from the end and clamp, matching `ecma:array.slice`
/// and `slices::emit_contiguous` — so a caller that already normalized its
/// bounds gets the same answer here as on the UTF-16 path.
pub fn emit_scalar_substring(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = chunks[current].alloc_scratch(2);
    let (start, end) = (base, base + 1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, end, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, start, line);

    emit_scalar_chars(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, start, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, end, line);
    crate::primitives::collections::emit_slice(chunks, current, line);
    chunks[current].emit_string_const("", line);
    crate::primitives::collections::emit_join(chunks, current, line);
}

/// Code-point index of a substring. Stack: `[haystack, needle]` → `[i32]`,
/// `-1` when absent. `mb_strpos("a😀b", "b")` is 2 where the UTF-16 index is 3.
///
/// Finds the UTF-16 index first and converts, rather than scanning code points:
/// the needle match itself is unit-independent, so only the PREFIX has to be
/// re-counted.
pub fn emit_scalar_index_of(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = chunks[current].alloc_scratch(2);
    let (hay, at) = (base, base + 1);

    // at = indexOf(hay, needle)  — a UTF-16 index
    {
        let chunk = &mut chunks[current];
        let needle = chunk.alloc_scratch(1);
        chunk.emit_op_u16(Op::LOCAL_SET, needle, line);
        chunk.emit_op_u16(Op::LOCAL_SET, hay, line);
        chunk.emit_op_u16(Op::LOCAL_GET, hay, line);
        chunk.emit_op_u16(Op::LOCAL_GET, needle, line);
        emit_index_of(chunk, line);
        chunk.emit_op_u16(Op::LOCAL_SET, at, line);

        // Absent stays absent — do not re-count a prefix of length -1.
        chunk.emit_op_u16(Op::LOCAL_GET, at, line);
        crate::primitives::instructions::core_wasm::i32_const(chunk, line, 0);
        crate::primitives::ops::emit_dyn_lt(chunk, line);
        crate::primitives::ops::emit_dyn_to_bool(chunk, line);
        chunk.emit_if_value(line);
        crate::primitives::instructions::core_wasm::i32_const(chunk, line, -1);
        chunk.emit_else(line);
        // Code points in hay[0..at].
        chunk.emit_op_u16(Op::LOCAL_GET, hay, line);
        crate::primitives::instructions::core_wasm::i32_const(chunk, line, 0);
        chunk.emit_op_u16(Op::LOCAL_GET, at, line);
        emit_substring(chunk, line);
    }
    emit_scalar_length(chunks, current, line);
    chunks[current].emit_end(line);
}

/// Code-point index of the LAST occurrence. Stack: `[haystack, needle]` →
/// `[i32]`, `-1` when absent. The scalar counterpart of `emit_last_index_of`.
pub fn emit_scalar_last_index_of(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = chunks[current].alloc_scratch(2);
    let (hay, at) = (base, base + 1);
    {
        let chunk = &mut chunks[current];
        let needle = chunk.alloc_scratch(1);
        chunk.emit_op_u16(Op::LOCAL_SET, needle, line);
        chunk.emit_op_u16(Op::LOCAL_SET, hay, line);
        chunk.emit_op_u16(Op::LOCAL_GET, hay, line);
        chunk.emit_op_u16(Op::LOCAL_GET, needle, line);
        emit_last_index_of(chunk, line);
        chunk.emit_op_u16(Op::LOCAL_SET, at, line);

        chunk.emit_op_u16(Op::LOCAL_GET, at, line);
        crate::primitives::instructions::core_wasm::i32_const(chunk, line, 0);
        crate::primitives::ops::emit_dyn_lt(chunk, line);
        crate::primitives::ops::emit_dyn_to_bool(chunk, line);
        chunk.emit_if_value(line);
        crate::primitives::instructions::core_wasm::i32_const(chunk, line, -1);
        chunk.emit_else(line);
        chunk.emit_op_u16(Op::LOCAL_GET, hay, line);
        crate::primitives::instructions::core_wasm::i32_const(chunk, line, 0);
        chunk.emit_op_u16(Op::LOCAL_GET, at, line);
        emit_substring(chunk, line);
    }
    emit_scalar_length(chunks, current, line);
    chunks[current].emit_end(line);
}

/// The code point at UTF-16 index `i`. Stack: `[string, i]` → `[i32]`.
/// Unlike `charCodeAt` this yields the whole astral code point rather than the
/// leading surrogate — `ord("\u{1F600}")` is 128512, not 55357.
pub fn emit_code_point_at(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("wasm:js-string", "codePointAt");
    chunk.emit_call(idx, 2, line);
}

/// The FIRST code point of a string. Stack: `[string]` → `[i32]`.
/// This is Python `ord`, and it is NOT `charCodeAt(s, 0)`: the two agree only
/// inside the BMP.
pub fn emit_first_code_point(chunk: &mut Chunk, line: u32) {
    crate::primitives::instructions::core_wasm::i32_const(chunk, line, 0);
    emit_code_point_at(chunk, line);
}

// ── Adapter primitive: trim with a character set ───────────────────────
//
// Local slot/const shorthands; this file otherwise emits raw ops.

fn get(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn set(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}

fn i32c(chunk: &mut Chunk, v: i32, line: u32) {
    chunk.emit_i32_const(v, line);
}

//
// `ecma:string.trim` / `trimStart` / `trimEnd` take NO character set — the
// ECMA definition is fixed whitespace. PHP `trim($s, $chars)`, Python
// `strip(chars)` / `lstrip(chars)` / `rstrip(chars)`, and Ruby `delete` all
// need one, so each language grew its own copy. This is that behaviour, once.
//
// Parameterized exactly like `primitives/slices.rs::Options`: front-ends build
// [`TrimOptions`] from profile properties, never from a language name.
//
// Stack: `[s]` or `[s, chars]` → `[string]`.

/// Which ends to trim, and what to strip when the caller passes no set.
#[derive(Clone, Copy)]
pub struct TrimOptions {
    pub left: bool,
    pub right: bool,
    /// Characters stripped when no explicit set is given.
    ///
    /// `None` delegates to `ecma:string.trim`/`trimStart`/`trimEnd` — the ECMA
    /// whitespace definition, which is what JS, Python, Dart and VB want.
    /// PHP passes `Some(" \t\n\r\0\x0B")`: its default includes NUL and
    /// vertical tab, which ECMA's does not.
    pub default_chars: Option<&'static str>,
}

impl TrimOptions {
    pub const fn both(default_chars: Option<&'static str>) -> TrimOptions {
        TrimOptions {
            left: true,
            right: true,
            default_chars,
        }
    }
    pub const fn start(default_chars: Option<&'static str>) -> TrimOptions {
        TrimOptions {
            left: true,
            right: false,
            default_chars,
        }
    }
    pub const fn end(default_chars: Option<&'static str>) -> TrimOptions {
        TrimOptions {
            left: false,
            right: true,
            default_chars,
        }
    }
}

/// Trim `opts.left`/`opts.right` ends of a string against a character set.
///
/// `argc` counts the values the caller left on the stack INCLUDING the string,
/// so `argc >= 2` means an explicit set was supplied. When it was not and
/// `opts.default_chars` is `None`, this delegates to the ECMA trim rather than
/// walking — the standard behaviour is the right behaviour there.
///
/// Walks by index and slices once at the end, rather than re-slicing per
/// character. Counts in UTF-16 units, matching every other helper in this file.
pub fn emit_trim_chars(
    chunks: &mut [Chunk],
    current: usize,
    argc: u8,
    opts: TrimOptions,
    line: u32,
) {
    if argc < 2 {
        match opts.default_chars {
            None => {
                // Standard behaviour — delegate.
                match (opts.left, opts.right) {
                    (true, true) => emit_trim(&mut chunks[current], line),
                    (true, false) => emit_trim_start(&mut chunks[current], line),
                    (false, true) => emit_trim_end(&mut chunks[current], line),
                    (false, false) => {}
                }
                return;
            }
            Some(defaults) => chunks[current].emit_string_const(defaults, line),
        }
    }

    let base = chunks[current].alloc_scratch(4);
    let (chars, s, start, end) = (base, base + 1, base + 2, base + 3);
    set(&mut chunks[current], chars, line);
    set(&mut chunks[current], s, line);

    i32c(&mut chunks[current], 0, line);
    set(&mut chunks[current], start, line);
    get(&mut chunks[current], s, line);
    emit_length(&mut chunks[current], line);
    set(&mut chunks[current], end, line);

    if opts.left {
        // while start < end && chars contains s[start] { start += 1 }
        let st = crate::primitives::loops::emit_loop_start(chunks, current, line);
        get(&mut chunks[current], start, line);
        get(&mut chunks[current], end, line);
        crate::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
        crate::primitives::loops::emit_loop_cond(chunks, current, line);
        emit_char_in_set(chunks, current, chars, s, start, 0, line);
        chunks[current].emit_br_if(st.break_depth(0) as u32, line);
        get(&mut chunks[current], start, line);
        i32c(&mut chunks[current], 1, line);
        chunks[current].emit_op(Op::I32_ADD, line);
        set(&mut chunks[current], start, line);
        crate::primitives::loops::emit_loop_end(chunks, current, st, line);
    }

    if opts.right {
        // while end > start && chars contains s[end-1] { end -= 1 }
        let st = crate::primitives::loops::emit_loop_start(chunks, current, line);
        get(&mut chunks[current], end, line);
        get(&mut chunks[current], start, line);
        crate::primitives::ops::emit_dyn_gt(&mut chunks[current], line);
        crate::primitives::loops::emit_loop_cond(chunks, current, line);
        emit_char_in_set(chunks, current, chars, s, end, -1, line);
        chunks[current].emit_br_if(st.break_depth(0) as u32, line);
        get(&mut chunks[current], end, line);
        i32c(&mut chunks[current], 1, line);
        chunks[current].emit_op(Op::I32_SUB, line);
        set(&mut chunks[current], end, line);
        crate::primitives::loops::emit_loop_end(chunks, current, st, line);
    }

    get(&mut chunks[current], s, line);
    get(&mut chunks[current], start, line);
    get(&mut chunks[current], end, line);
    emit_substring(&mut chunks[current], line);
}

/// Push i32 `1` when `chars` does NOT contain `s[idx + offset]` — the loop's
/// break condition. Uses `indexOf` on a one-character substring so a set
/// containing regex-significant characters is still matched literally.
fn emit_char_in_set(
    chunks: &mut [Chunk],
    current: usize,
    chars: u16,
    s: u16,
    idx: u16,
    offset: i32,
    line: u32,
) {
    let chunk = &mut chunks[current];
    get(chunk, chars, line);
    get(chunk, s, line);
    get(chunk, idx, line);
    if offset != 0 {
        i32c(chunk, -offset, line);
        chunk.emit_op(Op::I32_SUB, line);
    }
    let char_at = chunk.add_import("ecma:string", "charAt");
    chunk.emit_call(char_at, 2, line);
    let index_of = chunk.add_import("ecma:string", "indexOf");
    chunk.emit_call(index_of, 2, line);
    i32c(chunk, 0, line);
    crate::primitives::ops::emit_dyn_lt(chunk, line);
}
