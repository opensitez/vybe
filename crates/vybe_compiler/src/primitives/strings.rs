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
// The **`byte`** unit splits in two, and only half of it is blocked.
//
// Byte LENGTH is just a count, so it is here — see [`emit_byte_length`]. PHP
// `strlen`, Lua `#` and Go `len(s)` all want it, and until now the ONLY
// implementation on the platform was php's private one; `wasm:js-string.length`
// and `ecma:string.length` are both UTF-16 counts, so there was nothing else to
// reach for.
//
// Byte SUBSTRING is the blocked half: a byte offset can cut a code point in
// half, and `Value::String` (`Arc<str>`) cannot hold that. It stays out until
// `Literal::Bytes` (plan §3c) exists — not merely unwritten.

/// UTF-8 byte length. Stack: `[string]` → `[number]`.
///
/// `strlen("é")` is 2 and `strlen("😀")` is 4, where the UTF-16 count is 1 and
/// 2 and the code-point count is 1 and 1. Walks the UTF-16 units, reads each
/// whole code point, and adds its UTF-8 width — 1 below U+0080, 2 below U+0800,
/// 3 below U+10000, else 4 — stepping two units across an astral pair so the
/// low surrogate is never counted twice.
///
/// **No coercion here.** PHP's `strlen(123)` is `3` because php coerces first;
/// that belongs at the call site, not in the shared counter.
/// Truncate a C string at its first NUL. Stack: [string] → [string]
///
/// C strings are NUL-TERMINATED: the byte is a terminator, not content. Vybe
/// backs them with a JS string, where a `\0` is an ordinary code unit — so
/// `char u[8] = "ab"; u[1] = '\0';` left `strlen` reporting 2 and `strcmp`
/// comparing two characters, while `printf("%s")` correctly stopped at the
/// NUL. Measured against `cc`, which reports 1.
pub fn emit_cstr_truncate(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = chunks[current].alloc_scratch(2);
    let (s, at) = (base, base + 1);

    set(&mut chunks[current], s, line);
    get(&mut chunks[current], s, line);
    chunks[current].emit_string_const("\0", line);
    emit_index_of(&mut chunks[current], line);
    set(&mut chunks[current], at, line);

    // No NUL → the whole string is content.
    get(&mut chunks[current], at, line);
    chunks[current].emit_f64_const(0.0, line);
    crate::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
    crate::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], s, line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], s, line);
    chunks[current].emit_f64_const(0.0, line);
    get(&mut chunks[current], at, line);
    emit_substring(&mut chunks[current], line);
    chunks[current].emit_end(line);
}

/// `strlen` — bytes up to the first NUL. The fourth string-length UNIT,
/// alongside UTF-16 code units, code points and UTF-8 bytes; see
/// `builtin_slots.rs`. Reachable as `common:str_cstr_length`, so a profile can
/// declare it with `[builtin_slots.string] len` rather than reaching for an
/// emitter directly.
pub fn emit_cstr_length(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_cstr_truncate(chunks, current, line);
    emit_byte_length(chunks, current, line);
}

pub fn emit_byte_length(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = chunks[current].alloc_scratch(5);
    let (s, i, n, bytes, cp) = (base, base + 1, base + 2, base + 3, base + 4);

    set(&mut chunks[current], s, line);
    chunks[current].emit_f64_const(0.0, line);
    set(&mut chunks[current], i, line);
    chunks[current].emit_f64_const(0.0, line);
    set(&mut chunks[current], bytes, line);
    get(&mut chunks[current], s, line);
    emit_length(&mut chunks[current], line);
    set(&mut chunks[current], n, line);

    let loop_state = crate::primitives::loops::emit_loop_start(chunks, current, line);
    get(&mut chunks[current], i, line);
    get(&mut chunks[current], n, line);
    crate::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
    crate::primitives::loops::emit_loop_cond(chunks, current, line);

    // cp = s.codePointAt(i) — the WHOLE code point, not a surrogate half.
    get(&mut chunks[current], s, line);
    get(&mut chunks[current], i, line);
    {
        let idx = chunks[current].add_import("wasm:js-string", "codePointAt");
        chunks[current].emit_call(idx, 2, line);
    }
    set(&mut chunks[current], cp, line);

    // bytes += UTF-8 width of cp
    for (bound, width) in [(128.0, 1.0), (2048.0, 2.0), (65536.0, 3.0)] {
        get(&mut chunks[current], cp, line);
        chunks[current].emit_f64_const(bound, line);
        crate::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
        crate::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
        chunks[current].emit_if_value(line);
        chunks[current].emit_f64_const(width, line);
        chunks[current].emit_else(line);
    }
    chunks[current].emit_f64_const(4.0, line);
    for _ in 0..3 {
        chunks[current].emit_end(line);
    }
    get(&mut chunks[current], bytes, line);
    crate::primitives::ops::emit_dyn_add(&mut chunks[current], line);
    set(&mut chunks[current], bytes, line);

    // i += cp > 0xFFFF ? 2 : 1 — an astral code point spans two UTF-16 units.
    get(&mut chunks[current], cp, line);
    chunks[current].emit_f64_const(65535.0, line);
    crate::primitives::ops::emit_dyn_gt(&mut chunks[current], line);
    crate::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_f64_const(2.0, line);
    chunks[current].emit_else(line);
    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_end(line);
    get(&mut chunks[current], i, line);
    crate::primitives::ops::emit_dyn_add(&mut chunks[current], line);
    set(&mut chunks[current], i, line);

    crate::primitives::loops::emit_loop_end(chunks, current, loop_state, line);
    get(&mut chunks[current], bytes, line);
}

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

// ── split with a limit ──────────────────────────────────────────────────────
//
// `ecma:string.split` takes a limit, but it TRUNCATES: `"a,b,c".split(",", 2)`
// is `["a", "b"]` and the rest is thrown away. Every language that spells this
// keeps the remainder instead, re-joined into the final element — python
// `split(sep, maxsplit)`, php `explode($sep, $s, $limit)`, java
// `split(re, limit)`, go `SplitN`. So each one wrote the same "split fully,
// then re-join the tail" dance; python and php had it verbatim twice.
//
// They disagree on exactly TWO things, which is why this is one emitter with
// two flags rather than two implementations. Measured against real `python3`
// and `php` on `"a,b,c"`:
//
// | limit | py `split(",", n)` | php `explode(",", s, n)` |
// |---|---|---|
// | `0` | `["a,b,c"]` | `["a,b,c"]` |
// | `1` | `["a", "b,c"]` | `["a,b,c"]` |
// | `2` | `["a", "b", "c"]` | `["a", "b,c"]` |
// | `-1` | `["a", "b", "c"]` (unlimited) | `["a", "b"]` (drop the last) |
//
// python's limit counts SPLITS, php's counts PIECES — an off-by-one, not a
// different algorithm — and they read a negative limit differently.
//
// # What this does NOT cover yet
//
// Three limits, recorded so the next binding knows what it is walking into
// rather than discovering it as a red test:
//
// - **Left-to-right only.** python's `rsplit` consumes the limit from the END,
//   which is not this traversal; `emit_rsplit` in the python adapter still owns
//   it. Folding it in means a `from_end` flag AND re-joining the HEAD instead
//   of the tail, so it is a real change here, not a binding.
// - **Literal separators only.** This delegates to `ecma:string.split`, so a
//   REGEX separator — java `split(regex, limit)`, php `preg_split`, python
//   `re.split` — does not route here. Those are regex ops, out of scope per the
//   plan, but the shared emitter would need a `separator_is_pattern` flag.
// - **Two negative-limit policies, and java has a THIRD.** php drops trailing
//   pieces, python reads it as unlimited — and java's negative limit means
//   "keep trailing empty strings", which is a different axis again (it changes
//   whether empties SURVIVE, not how many pieces there are). Binding java will
//   need an enum here, not a bool.
//
// Python's no-separator form is also not here: `s.split()` splits on `\s+` runs
// AND drops leading/trailing empties, which is a different algorithm rather than
// a parameter, so it stays in the python adapter.

/// What a split limit counts, and what a negative one means.
#[derive(Clone, Copy)]
pub struct SplitOptions {
    /// The limit counts resulting PIECES (php `explode`) rather than the number
    /// of SPLITS performed (python `maxsplit`).
    pub limit_is_pieces: bool,
    /// A negative limit drops that many pieces off the END (php). Otherwise a
    /// negative limit means "no limit", which is python's `maxsplit=-1`.
    pub negative_drops_tail: bool }

impl SplitOptions {
    /// python `str.split(sep, maxsplit)`.
    pub const fn max_splits() -> SplitOptions {
        SplitOptions {
            limit_is_pieces: false,
            negative_drops_tail: false }
    }
    /// php `explode(separator, string, limit)`.
    pub const fn max_pieces() -> SplitOptions {
        SplitOptions {
            limit_is_pieces: true,
            negative_drops_tail: true }
    }
}

/// Split on a separator, keeping the remainder past the limit as the final
/// element. Stack: `[s, sep]` or `[s, sep, limit]` → `[array]`.
///
/// With no limit this is plain [`emit_split`] — the ECMA behaviour is the right
/// behaviour, so it delegates.
pub fn emit_split_limit(
    chunks: &mut [Chunk],
    current: usize,
    argc: u8,
    opts: SplitOptions,
    line: u32,
) {
    if argc < 3 {
        emit_split(&mut chunks[current], line);
        return;
    }

    let base = chunks[current].alloc_scratch(6);
    let (limit, sep, s, full, keep, tail) = (base, base + 1, base + 2, base + 3, base + 4, base + 5);

    set(&mut chunks[current], limit, line);
    set(&mut chunks[current], sep, line);
    set(&mut chunks[current], s, line);

    get(&mut chunks[current], s, line);
    get(&mut chunks[current], sep, line);
    emit_split(&mut chunks[current], line);
    set(&mut chunks[current], full, line);

    get(&mut chunks[current], limit, line);
    chunks[current].emit_f64_const(0.0, line);
    super::ops::emit_dyn_lt(&mut chunks[current], line);
    super::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    if opts.negative_drops_tail {
        // `full.slice(0, max(full.length + limit, 0))`.
        get(&mut chunks[current], full, line);
        chunks[current].emit_i32_const(0, line);
        get(&mut chunks[current], full, line);
        chunks[current].emit_op(Op::ARRAY_LENGTH, line);
        get(&mut chunks[current], limit, line);
        chunks[current].emit_op(Op::F64_ADD, line);
        emit_clamp_low_zero(chunks, current, line);
        array_call(chunks, current, "slice", 3, line);
    } else {
        get(&mut chunks[current], full, line);
    }
    chunks[current].emit_else(line);
    {
        // How many pieces stay whole before the remainder is re-joined.
        get(&mut chunks[current], limit, line);
        if opts.limit_is_pieces {
            chunks[current].emit_f64_const(1.0, line);
            chunks[current].emit_op(Op::F64_SUB, line);
            emit_clamp_low_zero(chunks, current, line);
        }
        set(&mut chunks[current], keep, line);

        get(&mut chunks[current], full, line);
        get(&mut chunks[current], keep, line);
        chunks[current].emit_i32_const(0x7FFF_FFFF, line);
        array_call(chunks, current, "slice", 3, line);
        set(&mut chunks[current], tail, line);

        get(&mut chunks[current], full, line);
        chunks[current].emit_i32_const(0, line);
        get(&mut chunks[current], keep, line);
        array_call(chunks, current, "slice", 3, line);
        // `head` reuses `full`: the fully-split array is no longer needed.
        set(&mut chunks[current], full, line);

        get(&mut chunks[current], tail, line);
        chunks[current].emit_op(Op::ARRAY_LENGTH, line);
        super::ops::emit_dyn_to_bool(&mut chunks[current], line);
        chunks[current].emit_if(line);
        get(&mut chunks[current], full, line);
        get(&mut chunks[current], tail, line);
        get(&mut chunks[current], sep, line);
        array_call(chunks, current, "join", 2, line);
        array_call(chunks, current, "push", 2, line);
        chunks[current].emit_op(Op::DROP, line);
        chunks[current].emit_end(line);

        get(&mut chunks[current], full, line);
    }
    chunks[current].emit_end(line);
}

/// `max(n, 0)`. Stack: `[n]` → `[n]`.
fn emit_clamp_low_zero(chunks: &mut [Chunk], current: usize, line: u32) {
    let n = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], n, line);
    get(&mut chunks[current], n, line);
    chunks[current].emit_f64_const(0.0, line);
    super::ops::emit_dyn_lt(&mut chunks[current], line);
    super::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], n, line);
    chunks[current].emit_end(line);
}

fn array_call(chunks: &mut [Chunk], current: usize, name: &str, argc: u8, line: u32) {
    let idx = chunks[current].add_import("ecma:array", name.to_string());
    chunks[current].emit_call(idx, argc, line);
}

// ── pad to a width ──────────────────────────────────────────────────────────
//
// `ecma:string.padStart`/`padEnd` (ECMA-262 §22.1.3.16) already do one-sided
// padding exactly as every language wants it, including the two guards each
// implementation was re-checking by hand: `StringPad` returns the input
// unchanged when the target is not longer than the string, and when the filler
// is the empty string. So one-sided padding DELEGATES and needs nothing here.
//
// CENTRED padding is the part ECMA has no operation for, and it was written
// twice — php `str_pad(..., STR_PAD_BOTH)` and python `str.center`. They agree
// everywhere except which side gets the odd character. Measured against real
// `php` and `python3`:
//
// | `"ab"`, width | py `center` | php `STR_PAD_BOTH` |
// |---|---|---|
// | 5 | `"  ab "` | `"_ab__"` |
// | 6 | `"  ab  "` | `"__ab__"` |
// | 7 | `"   ab  "` | `"__ab___"` |
// | `"abc"`, 8 | `"--abc---"` | `"--abc---"` |
//
// php always leaves the extra on the right. CPython's rule is
// `left = marg/2 + (marg & width & 1)` — the extra moves LEFT only when the
// margin and the width are both odd, which is why width 8 agrees and 5 and 7
// do not. That is [`CenterBias`], and it is the whole of the difference.

/// Which side (or sides) of a string the padding goes on.
#[derive(Clone, Copy, PartialEq, Eq)]
pub enum PadSide {
    /// Right-align — python `rjust`, php `STR_PAD_LEFT`.
    Start,
    /// Left-align — python `ljust`, php `STR_PAD_RIGHT`.
    End,
    /// Centre — python `center`, php `STR_PAD_BOTH`.
    Both }

/// Where the odd character goes when a centred pad does not divide evenly.
#[derive(Clone, Copy, PartialEq, Eq)]
pub enum CenterBias {
    /// Always on the right — php `STR_PAD_BOTH`.
    Right,
    /// On the left when the margin AND the width are both odd, on the right
    /// otherwise — CPython `str.center`.
    LeftWhenBothOdd }

/// Pad `s` out to `width`. Stack: `[s, width]` or `[s, width, fill]` → `[string]`.
///
/// `argc` counts the values on the stack INCLUDING the string, so `argc >= 3`
/// means an explicit filler was supplied; otherwise a space is used.
pub fn emit_pad(
    chunks: &mut [Chunk],
    current: usize,
    argc: u8,
    side: PadSide,
    bias: CenterBias,
    line: u32,
) {
    if argc < 3 {
        chunks[current].emit_string_const(" ", line);
    }
    match side {
        PadSide::Start => pad_call(chunks, current, "padStart", line),
        PadSide::End => pad_call(chunks, current, "padEnd", line),
        PadSide::Both => emit_pad_both(chunks, current, bias, line) }
}

fn pad_call(chunks: &mut [Chunk], current: usize, name: &str, line: u32) {
    let idx = chunks[current].add_import("ecma:string", name.to_string());
    chunks[current].emit_call(idx, 3, line);
}

/// Centre `s` in `width`: pad the left to the split point, then the right to
/// the full width. Both steps are `StringPad`, so a width at or below the
/// string's own length is a no-op without a guard here.
fn emit_pad_both(chunks: &mut [Chunk], current: usize, bias: CenterBias, line: u32) {
    let base = chunks[current].alloc_scratch(4);
    let (fill, width, s, len) = (base, base + 1, base + 2, base + 3);

    set(&mut chunks[current], fill, line);
    super::convert::emit_to_int(&mut chunks[current], line);
    set(&mut chunks[current], width, line);
    set(&mut chunks[current], s, line);

    get(&mut chunks[current], s, line);
    emit_length(&mut chunks[current], line);
    set(&mut chunks[current], len, line);

    // left_target = len + margin/2 [+ 1 when the bias says so]
    get(&mut chunks[current], s, line);
    get(&mut chunks[current], len, line);
    emit_margin(chunks, current, width, len, line);
    i32c(&mut chunks[current], 2, line);
    chunks[current].emit_op(Op::I32_DIV_S, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    if bias == CenterBias::LeftWhenBothOdd {
        emit_margin(chunks, current, width, len, line);
        get(&mut chunks[current], width, line);
        chunks[current].emit_op(Op::I32_AND, line);
        i32c(&mut chunks[current], 1, line);
        chunks[current].emit_op(Op::I32_AND, line);
        chunks[current].emit_op(Op::I32_ADD, line);
    }
    get(&mut chunks[current], fill, line);
    pad_call(chunks, current, "padStart", line);

    get(&mut chunks[current], width, line);
    get(&mut chunks[current], fill, line);
    pad_call(chunks, current, "padEnd", line);
}

/// `width - len`. Stack: `-> [i32]`.
fn emit_margin(chunks: &mut [Chunk], current: usize, width: u16, len: u16, line: u32) {
    get(&mut chunks[current], width, line);
    get(&mut chunks[current], len, line);
    chunks[current].emit_op(Op::I32_SUB, line);
}

// ── glob matching ───────────────────────────────────────────────────────────
//
// Shell-style patterns — `*`, `?` — are php `fnmatch`, python
// `fnmatch.fnmatch`/`fnmatchcase`, Go `path.Match`, Ruby `File.fnmatch`. They
// were written twice on this platform, in different SHAPES: php as an emitter
// that rewrites the glob to a regex, python as a hand-written iterative matcher
// inside a substring-gated Python-source PRELUDE.
//
// Both reduce to "escape the regex metacharacters, translate `*` and `?`,
// anchor". The one behavioural difference is case: python's `fnmatch` folds
// both sides and `fnmatchcase` does not, while php's never folds. That is
// [`GlobOptions::fold_case`].
//
// Neither implementation supported `[seq]` character classes, which real
// `fnmatch` has — so this does not either. Recorded rather than silently
// carried forward: adding it is a change in behaviour for both languages and
// wants measuring against the real runtimes first.

/// Per-language glob rules.
#[derive(Clone, Copy)]
pub struct GlobOptions {
    /// Lower-case both pattern and subject before matching — python
    /// `fnmatch.fnmatch`. php `fnmatch` and python `fnmatchcase` do not.
    pub fold_case: bool }

impl GlobOptions {
    /// php `fnmatch`, python `fnmatchcase`.
    pub const fn exact() -> GlobOptions {
        GlobOptions { fold_case: false }
    }
    /// python `fnmatch.fnmatch`.
    pub const fn folded() -> GlobOptions {
        GlobOptions { fold_case: true }
    }
}

/// The regex metacharacters a glob escapes OUTSIDE a character class. `*`, `?`
/// and `[` are absent — they are the glob operators, handled by the scan.
const GLOB_META: &str = "\\.()+^${}|";

/// Translate a glob to an anchored regex SOURCE. Stack: `[pattern]` → `[string]`.
///
/// This is python `fnmatch.translate` on its own, and the first half of every
/// glob match.
///
/// A SCAN rather than a chain of `replaceAll` calls, because `[seq]` classes
/// make the escaping context-dependent: `.` is a metacharacter outside a class
/// and a literal inside one, and `[`/`]` delimit rather than escape. Both
/// previous implementations used the chain, escaped `[` and `]`, and so matched
/// `gr[ae]y` LITERALLY — `fnmatch("*gr[ae]y", "color_is_grey")` was false in
/// both php and python where both real runtimes say true.
pub fn emit_glob_to_regex(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = chunks[current].alloc_scratch(6);
    let (pat, out, i, n, in_class, c) = (base, base + 1, base + 2, base + 3, base + 4, base + 5);

    set(&mut chunks[current], pat, line);
    chunks[current].emit_string_const("^", line);
    set(&mut chunks[current], out, line);
    i32c(&mut chunks[current], 0, line);
    set(&mut chunks[current], i, line);
    i32c(&mut chunks[current], 0, line);
    set(&mut chunks[current], in_class, line);
    get(&mut chunks[current], pat, line);
    emit_length(&mut chunks[current], line);
    set(&mut chunks[current], n, line);

    let loop_state = crate::primitives::loops::emit_loop_start(chunks, current, line);
    get(&mut chunks[current], i, line);
    get(&mut chunks[current], n, line);
    crate::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
    crate::primitives::loops::emit_loop_cond(chunks, current, line);

    // c = pat[i]
    emit_char_at(chunks, current, pat, i, line);
    set(&mut chunks[current], c, line);

    get(&mut chunks[current], in_class, line);
    chunks[current].emit_if_value(line);
    {
        // Inside `[...]` everything is literal until the closing `]`.
        get(&mut chunks[current], c, line);
        chunks[current].emit_string_const("]", line);
        chunks[current].emit_op(Op::EQ, line);
        chunks[current].emit_if(line);
        i32c(&mut chunks[current], 0, line);
        set(&mut chunks[current], in_class, line);
        chunks[current].emit_end(line);
        get(&mut chunks[current], c, line);
    }
    chunks[current].emit_else(line);
    {
        get(&mut chunks[current], c, line);
        chunks[current].emit_string_const("*", line);
        chunks[current].emit_op(Op::EQ, line);
        chunks[current].emit_if_value(line);
        chunks[current].emit_string_const(".*", line);
        chunks[current].emit_else(line);

        get(&mut chunks[current], c, line);
        chunks[current].emit_string_const("?", line);
        chunks[current].emit_op(Op::EQ, line);
        chunks[current].emit_if_value(line);
        chunks[current].emit_string_const(".", line);
        chunks[current].emit_else(line);

        get(&mut chunks[current], c, line);
        chunks[current].emit_string_const("[", line);
        chunks[current].emit_op(Op::EQ, line);
        chunks[current].emit_if_value(line);
        {
            i32c(&mut chunks[current], 1, line);
            set(&mut chunks[current], in_class, line);
            // A leading `!` negates in glob; regex spells that `^`. Consume it
            // here so the scan does not re-emit it as a literal.
            let nxt = chunks[current].alloc_scratch(1);
            get(&mut chunks[current], i, line);
            i32c(&mut chunks[current], 1, line);
            chunks[current].emit_op(Op::I32_ADD, line);
            set(&mut chunks[current], nxt, line);
            emit_char_at(chunks, current, pat, nxt, line);
            chunks[current].emit_string_const("!", line);
            chunks[current].emit_op(Op::EQ, line);
            chunks[current].emit_if_value(line);
            get(&mut chunks[current], nxt, line);
            set(&mut chunks[current], i, line);
            chunks[current].emit_string_const("[^", line);
            chunks[current].emit_else(line);
            chunks[current].emit_string_const("[", line);
            chunks[current].emit_end(line);
        }
        chunks[current].emit_else(line);
        {
            // Escape only where the character is a regex metacharacter.
            chunks[current].emit_string_const(GLOB_META, line);
            get(&mut chunks[current], c, line);
            emit_index_of(&mut chunks[current], line);
            i32c(&mut chunks[current], 0, line);
            crate::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
            crate::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
            chunks[current].emit_if_value(line);
            get(&mut chunks[current], c, line);
            chunks[current].emit_else(line);
            chunks[current].emit_string_const("\\", line);
            get(&mut chunks[current], c, line);
            concat(chunks, current, line);
            chunks[current].emit_end(line);
        }
        chunks[current].emit_end(line);
        chunks[current].emit_end(line);
        chunks[current].emit_end(line);
    }
    chunks[current].emit_end(line);

    // out += <the piece just computed>
    let piece = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], piece, line);
    get(&mut chunks[current], out, line);
    get(&mut chunks[current], piece, line);
    concat(chunks, current, line);
    set(&mut chunks[current], out, line);

    get(&mut chunks[current], i, line);
    i32c(&mut chunks[current], 1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    set(&mut chunks[current], i, line);
    crate::primitives::loops::emit_loop_end(chunks, current, loop_state, line);

    get(&mut chunks[current], out, line);
    chunks[current].emit_string_const("$", line);
    concat(chunks, current, line);
}

/// One character of `s` at `idx`, as a string. Stack: `-> [string]`.
fn emit_char_at(chunks: &mut [Chunk], current: usize, s: u16, idx: u16, line: u32) {
    get(&mut chunks[current], s, line);
    get(&mut chunks[current], idx, line);
    get(&mut chunks[current], idx, line);
    i32c(&mut chunks[current], 1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    emit_substring(&mut chunks[current], line);
}

/// Match `name` against a shell-style `pattern`, with an explicit REGEX FLAGS
/// string. Stack: `[name, pattern, flags]` → `[bool]`.
///
/// Flags are a stack value because php's `FNM_CASEFOLD` is a runtime argument —
/// `fnmatch($pat, $s, $flags)` — so the fold decision cannot be compile-time for
/// every caller. Python's two spellings ARE compile-time and use
/// [`emit_glob_match`], which pushes the constant for them.
///
/// Case-insensitivity is the regex `i` flag rather than lower-casing both
/// sides. Lower-casing is what python's prelude did, and it is wrong for
/// classes: `[A-Z]` becomes `[a-z]` and stops meaning what it said.
pub fn emit_glob_match_flagged(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = chunks[current].alloc_scratch(3);
    let (flags, pattern, name) = (base, base + 1, base + 2);
    set(&mut chunks[current], flags, line);
    set(&mut chunks[current], pattern, line);
    set(&mut chunks[current], name, line);

    get(&mut chunks[current], pattern, line);
    emit_glob_to_regex(chunks, current, line);
    get(&mut chunks[current], flags, line);
    let new_re = chunks[current].add_import("ecma:regexp", "new");
    chunks[current].emit_call(new_re, 2, line);

    get(&mut chunks[current], name, line);
    let exec = chunks[current].add_import("ecma:regexp", "exec");
    chunks[current].emit_call(exec, 2, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    crate::primitives::ops::emit_dyn_ne(&mut chunks[current], line);
    // `emit_dyn_ne` leaves an i32, and the VM has no true/false — a real
    // boolean is `wasm:js-boolean.fromI32`. php never noticed (its `?:` coerces),
    // but python `print` shows the TYPE, so an i32 rendered as `1`/`0` where
    // real python3 prints `True`/`False`.
    let from_i32 = chunks[current].add_import("wasm:js-boolean", "fromI32");
    chunks[current].emit_call(from_i32, 1, line);
}

/// Match `name` against a shell-style `pattern`.
/// Stack: `[name, pattern]` → `[bool]`.
pub fn emit_glob_match(chunks: &mut [Chunk], current: usize, opts: GlobOptions, line: u32) {
    chunks[current].emit_string_const(if opts.fold_case { "i" } else { "" }, line);
    emit_glob_match_flagged(chunks, current, line);
}

fn concat(chunks: &mut [Chunk], current: usize, line: u32) {
    let idx = chunks[current].add_import("wasm:js-string", "concat");
    chunks[current].emit_call(idx, 2, line);
}

// ── digit grouping ──────────────────────────────────────────────────────────
//
// Thousands separators over an ALREADY-FORMATTED numeric string. Written three
// times on this platform: php `number_format`, python's `__py_fmt_group`
// (the `,` format-spec flag), and java's `%,d` inside an AST prelude. `printf`
// has no grouping in any of them, so each did the same string surgery.
//
// The separators are STACK values, not options: php takes them as runtime
// arguments (`number_format($n, $d, $dec, $thou)`) while python's are fixed
// `,` and `.`. A compile-time option could not serve php at all.
//
// Grouping is by THREE from the right, which is what all three implemented.
// Indian 2-3-2 grouping and locale-aware separators are ECMA-402
// (`Intl.NumberFormat`) and are not this.

/// Insert `group_sep` every three digits into the integer part, and use
/// `dec_point` for the fraction.
/// Stack: `[formatted, group_sep, dec_point]` → `[string]`.
///
/// `formatted` is a plain numeric string — `"-1234.5"`, `"1234"`. Any sign is
/// preserved and never grouped; the fractional part is copied through with its
/// separator replaced.
pub fn emit_group_digits(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = chunks[current].alloc_scratch(9);
    let (dec_point, sep, s, dot, int_end, digits, frac, out, i) = (
        base,
        base + 1,
        base + 2,
        base + 3,
        base + 4,
        base + 5,
        base + 6,
        base + 7,
        base + 8,
    );
    set(&mut chunks[current], dec_point, line);
    set(&mut chunks[current], sep, line);
    set(&mut chunks[current], s, line);

    // dot = s.indexOf(".");  int_end = dot < 0 ? s.length : dot
    get(&mut chunks[current], s, line);
    chunks[current].emit_string_const(".", line);
    emit_index_of(&mut chunks[current], line);
    set(&mut chunks[current], dot, line);
    get(&mut chunks[current], dot, line);
    i32c(&mut chunks[current], 0, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], s, line);
    emit_length(&mut chunks[current], line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], dot, line);
    chunks[current].emit_end(line);
    set(&mut chunks[current], int_end, line);

    // frac = dot < 0 ? "" : dec_point ++ s.slice(dot + 1)
    get(&mut chunks[current], dot, line);
    i32c(&mut chunks[current], 0, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_string_const("", line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], dec_point, line);
    get(&mut chunks[current], s, line);
    get(&mut chunks[current], dot, line);
    i32c(&mut chunks[current], 1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    i32c(&mut chunks[current], 0x7FFF_FFFF, line);
    emit_substring(&mut chunks[current], line);
    concat(chunks, current, line);
    chunks[current].emit_end(line);
    set(&mut chunks[current], frac, line);

    // A leading sign is copied out and never grouped.
    let start = chunks[current].alloc_scratch(1);
    get(&mut chunks[current], s, line);
    i32c(&mut chunks[current], 0, line);
    i32c(&mut chunks[current], 1, line);
    emit_substring(&mut chunks[current], line);
    let first = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], first, line);
    chunks[current].emit_string_const("+-", line);
    get(&mut chunks[current], first, line);
    emit_index_of(&mut chunks[current], line);
    i32c(&mut chunks[current], 0, line);
    chunks[current].emit_op(Op::I32_GE_S, line);
    // `indexOf("")` is 0, so an empty string would read as a sign — guard on
    // the character being non-empty.
    get(&mut chunks[current], first, line);
    emit_length(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_AND, line);
    chunks[current].emit_if_value(line);
    i32c(&mut chunks[current], 1, line);
    chunks[current].emit_else(line);
    i32c(&mut chunks[current], 0, line);
    chunks[current].emit_end(line);
    set(&mut chunks[current], start, line);

    get(&mut chunks[current], s, line);
    get(&mut chunks[current], start, line);
    get(&mut chunks[current], int_end, line);
    emit_substring(&mut chunks[current], line);
    set(&mut chunks[current], digits, line);

    get(&mut chunks[current], s, line);
    i32c(&mut chunks[current], 0, line);
    get(&mut chunks[current], start, line);
    emit_substring(&mut chunks[current], line);
    set(&mut chunks[current], out, line);

    // for i in 0..len(digits): if i > 0 && (len - i) % 3 == 0 { out += sep }
    let n = chunks[current].alloc_scratch(1);
    get(&mut chunks[current], digits, line);
    emit_length(&mut chunks[current], line);
    set(&mut chunks[current], n, line);
    i32c(&mut chunks[current], 0, line);
    set(&mut chunks[current], i, line);

    let loop_state = crate::primitives::loops::emit_loop_start(chunks, current, line);
    get(&mut chunks[current], i, line);
    get(&mut chunks[current], n, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    crate::primitives::loops::emit_loop_cond(chunks, current, line);

    get(&mut chunks[current], i, line);
    i32c(&mut chunks[current], 0, line);
    chunks[current].emit_op(Op::I32_GT_S, line);
    get(&mut chunks[current], n, line);
    get(&mut chunks[current], i, line);
    chunks[current].emit_op(Op::I32_SUB, line);
    i32c(&mut chunks[current], 3, line);
    chunks[current].emit_op(Op::I32_REM_S, line);
    i32c(&mut chunks[current], 0, line);
    chunks[current].emit_op(Op::I32_EQ, line);
    chunks[current].emit_op(Op::I32_AND, line);
    chunks[current].emit_if(line);
    get(&mut chunks[current], out, line);
    get(&mut chunks[current], sep, line);
    concat(chunks, current, line);
    set(&mut chunks[current], out, line);
    chunks[current].emit_end(line);

    get(&mut chunks[current], out, line);
    get(&mut chunks[current], digits, line);
    get(&mut chunks[current], i, line);
    get(&mut chunks[current], i, line);
    i32c(&mut chunks[current], 1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    emit_substring(&mut chunks[current], line);
    concat(chunks, current, line);
    set(&mut chunks[current], out, line);

    get(&mut chunks[current], i, line);
    i32c(&mut chunks[current], 1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    set(&mut chunks[current], i, line);
    crate::primitives::loops::emit_loop_end(chunks, current, loop_state, line);

    get(&mut chunks[current], out, line);
    get(&mut chunks[current], frac, line);
    concat(chunks, current, line);
}

// ── splice ──────────────────────────────────────────────────────────────────
//
// Remove a run of characters and/or insert text at a position — the operation
// behind pascal `Delete(var S; Index; Count)` and `Insert(Src; var Dst; Index)`,
// VB `Mid$` replacement, and JS `String` splicing. Both pascal procedures are
// the SAME splice with different arguments defaulted, which is why they share
// an emitter rather than getting one global each.
//
// Previously these were runtime-helper GLOBALS (`__vybe_pascal_str_insert`,
// `__vybe_pascal_str_remove_range`) reached by name through a bundle table.
// That indirection is what let `pascal.str_insert` point at `__vybe_str_insert`
// — a different helper with a different argument order — undetected.

/// Splice a string: drop `count` characters at `index` and put `insert` there.
/// Stack: `[s, index, count, insert]` → `[string]`. `index` is ZERO-based;
/// a language with 1-based string positions subtracts before calling.
pub fn emit_splice(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = chunks[current].alloc_scratch(4);
    let (insert, count, index, s) = (base, base + 1, base + 2, base + 3);
    set(&mut chunks[current], insert, line);
    set(&mut chunks[current], count, line);
    set(&mut chunks[current], index, line);
    set(&mut chunks[current], s, line);

    // s[0..index] ++ insert ++ s[index+count..]
    get(&mut chunks[current], s, line);
    i32c(&mut chunks[current], 0, line);
    get(&mut chunks[current], index, line);
    emit_substring(&mut chunks[current], line);

    get(&mut chunks[current], insert, line);
    concat(chunks, current, line);

    get(&mut chunks[current], s, line);
    get(&mut chunks[current], index, line);
    get(&mut chunks[current], count, line);
    crate::primitives::ops::emit_dyn_add(&mut chunks[current], line);
    i32c(&mut chunks[current], 0x7FFF_FFFF, line);
    emit_substring(&mut chunks[current], line);
    concat(chunks, current, line);
}

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
    pub default_chars: Option<&'static str> }

impl TrimOptions {
    pub const fn both(default_chars: Option<&'static str>) -> TrimOptions {
        TrimOptions {
            left: true,
            right: true,
            default_chars }
    }
    pub const fn start(default_chars: Option<&'static str>) -> TrimOptions {
        TrimOptions {
            left: true,
            right: false,
            default_chars }
    }
    pub const fn end(default_chars: Option<&'static str>) -> TrimOptions {
        TrimOptions {
            left: false,
            right: true,
            default_chars }
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
            Some(defaults) => chunks[current].emit_string_const(defaults, line) }
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

/// `sep.join(iterable)` — join ANY iterable, not only an array.
///
/// Most languages spell a join over a general iterable: Python's
/// `sep.join(x)`, PHP's `implode`, Go's `strings.Join`, C#'s `string.Join`,
/// Lua's `table.concat`. `ecma:array.join` is right to refuse a generator —
/// one has no `length`, and node returns `""` for
/// `Array.prototype.join.call(gen, sep)` exactly as vybex does. That is ECMA
/// behaviour and is left alone; accepting the iterable is the LANGUAGE's rule,
/// so it lives here as an adapter over the ECMA call rather than inside it.
///
/// Materialisation goes through `collections::emit_spread_iterable`, the shared
/// helper spread and destructuring already use: it drains a generator through
/// `generators.rs` stack-switching and everything else through the ECMA-262
/// iterator protocol. `ecma:array.from` is NOT enough — it leaves a Vybe
/// generator empty (verified with `vybex -d`).
///
/// Stack: `[iterable, separator]` → `[string]`.
pub fn emit_join_iterable(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let sep = chunk.alloc_scratch(2);
    let iterable = sep + 1;
    chunk.emit_op_u16(Op::LOCAL_SET, sep, line);
    chunk.emit_op_u16(Op::LOCAL_SET, iterable, line);

    chunk.emit_op_u16(Op::LOCAL_GET, iterable, line);
    crate::primitives::collections::emit_spread_iterable(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, sep, line);
    call_host(&mut chunks[current], "ecma:array", "join", 2, line);
}

/// Register the import on the CURRENT chunk so `normalize_import_table` remaps
/// it through that chunk's own table.
fn call_host(chunk: &mut Chunk, module: &str, name: &str, argc: u8, line: u32) {
    let idx = chunk.add_import(module, name);
    chunk.emit_call(idx, argc, line);
}
