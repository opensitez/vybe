//! `System.Net.WebUtility` — bytecode-only.
//!
//! The type was absent from the catalog, so `WebUtility.HtmlEncode(s)` trapped
//! with `undefined is not callable`.
//!
//! ## What .NET actually does, measured against `dotnet` 10
//!
//! `HtmlEncode` is NOT "escape the five specials". It escapes
//! `<`, `>`, `&`, `"` and `'` (as `&#39;`, not `&apos;`), **plus every code
//! point in `[160, 255]` as a decimal reference**, **plus every ASTRAL code
//! point as a decimal reference of the WHOLE code point**. Everything else is
//! left alone — including BMP characters above 255:
//!
//! ```text
//!   "<div id='1'>"  -> "&lt;div id=&#39;1&#39;&gt;"
//!   "café"          -> "caf&#233;"
//!   U+009F          -> unchanged        U+00A0 -> &#160;
//!   U+00FF -> &#255;                    U+0100 -> unchanged
//!   "€"  (U+20AC)   -> unchanged        — BMP, so not escaped
//!   "日本語"         -> unchanged
//!   "😀" (U+1F600)  -> "&#128512;"      — astral, so escaped as one code point
//! ```
//!
//! ⛔ The corpus only checks `HtmlDecode(HtmlEncode(s)) == s` on
//! `<div id='N'>`, which any SELF-INVERSE pair satisfies — including
//! "escape nothing at all". The measured encodings above are the real gate.
//!
//! `HtmlDecode` is not the encoder run backwards: it also reads hex references
//! and named entities the encoder never emits (`&#xE9;`, `&eacute;`), leaves an
//! unknown entity alone, and is SINGLE-PASS (`&amp;amp;` → `&amp;`).
//!
//! `UrlEncode` is `encodeURIComponent` plus three fixups. Its unreserved set is
//! `A-Za-z0-9-._!*()` with space as `+`, so relative to ECMA-262 §19.2.6.5 it
//! escapes `~` and `'` and folds `%20`:
//!
//! ```text
//!   "a(b)c" -> "a(b)c"      "a!b" -> "a!b"      "a*b" -> "a*b"
//!   "a~b"   -> "a%7Eb"      "a'b" -> "a%27b"    " x"  -> "+x"
//!   "café"  -> "caf%C3%A9"  — UTF-8 bytes, uppercase hex
//! ```
//!
//! Both HTML codecs are emitted as ONE deduplicated function chunk per module
//! rather than inline, so a program with many calls pays for the loop once.

use std::sync::Arc;

use vybe_compiler::primitives::ops;
use vybe_compiler::primitives::url::{self, PercentOptions};
use vybe_runtime::opcode::Op;
use vybe_runtime::{Chunk, Value};

/// .NET's `UrlEncode` differs from `encodeURIComponent` on exactly `~`, `'`
/// and the space; `PercentOptions::form()` already answers the first and the
/// third, so only the apostrophe is left to this adapter.
const URL_ENCODE_OPTIONS: PercentOptions = PercentOptions {
    space_as_plus: true,
    escape_tilde: true,
    safe: "",
};

/// The named entities `HtmlDecode` resolves. The encoder emits none of them
/// beyond the five specials — they are here because .NET decodes them, and a
/// decoder that only reverses its own encoder is not `HtmlDecode`.
///
/// ⛔ NOT the full HTML5 set (~2200 names): an unlisted entity is left
/// untouched, which is what .NET does for a genuinely unknown one — so the
/// shortfall is silent. Named here so the next reader does not re-measure it.
const NAMED_ENTITIES: &[(&str, &str)] = &[
    ("lt", "<"),
    ("gt", ">"),
    ("amp", "&"),
    ("quot", "\""),
    ("apos", "'"),
    ("nbsp", "\u{A0}"),
    ("copy", "\u{A9}"),
    ("reg", "\u{AE}"),
    ("trade", "\u{2122}"),
    ("euro", "\u{20AC}"),
    ("pound", "\u{A3}"),
    ("yen", "\u{A5}"),
    ("cent", "\u{A2}"),
    ("sect", "\u{A7}"),
    ("deg", "\u{B0}"),
    ("plusmn", "\u{B1}"),
    ("micro", "\u{B5}"),
    ("middot", "\u{B7}"),
    ("times", "\u{D7}"),
    ("divide", "\u{F7}"),
    ("laquo", "\u{AB}"),
    ("raquo", "\u{BB}"),
    ("hellip", "\u{2026}"),
    ("ndash", "\u{2013}"),
    ("mdash", "\u{2014}"),
    ("lsquo", "\u{2018}"),
    ("rsquo", "\u{2019}"),
    ("ldquo", "\u{201C}"),
    ("rdquo", "\u{201D}"),
    ("eacute", "\u{E9}"),
    ("egrave", "\u{E8}"),
    ("ecirc", "\u{EA}"),
    ("euml", "\u{EB}"),
    ("aacute", "\u{E1}"),
    ("agrave", "\u{E0}"),
    ("acirc", "\u{E2}"),
    ("auml", "\u{E4}"),
    ("aring", "\u{E5}"),
    ("aelig", "\u{E6}"),
    ("ccedil", "\u{E7}"),
    ("iacute", "\u{ED}"),
    ("igrave", "\u{EC}"),
    ("icirc", "\u{EE}"),
    ("iuml", "\u{EF}"),
    ("ntilde", "\u{F1}"),
    ("oacute", "\u{F3}"),
    ("ograve", "\u{F2}"),
    ("ocirc", "\u{F4}"),
    ("ouml", "\u{F6}"),
    ("oslash", "\u{F8}"),
    ("uacute", "\u{FA}"),
    ("ugrave", "\u{F9}"),
    ("ucirc", "\u{FB}"),
    ("uuml", "\u{FC}"),
    ("yacute", "\u{FD}"),
    ("yuml", "\u{FF}"),
    ("szlig", "\u{DF}"),
    ("iexcl", "\u{A1}"),
    ("iquest", "\u{BF}"),
    ("curren", "\u{A4}"),
    ("brvbar", "\u{A6}"),
    ("uml", "\u{A8}"),
    ("ordf", "\u{AA}"),
    ("not", "\u{AC}"),
    ("shy", "\u{AD}"),
    ("macr", "\u{AF}"),
    ("sup2", "\u{B2}"),
    ("sup3", "\u{B3}"),
    ("acute", "\u{B4}"),
    ("para", "\u{B6}"),
    ("cedil", "\u{B8}"),
    ("sup1", "\u{B9}"),
    ("ordm", "\u{BA}"),
    ("frac14", "\u{BC}"),
    ("frac12", "\u{BD}"),
    ("frac34", "\u{BE}"),
    ("bull", "\u{2022}"),
    ("dagger", "\u{2020}"),
    ("permil", "\u{2030}"),
    ("lsaquo", "\u{2039}"),
    ("rsaquo", "\u{203A}"),
    ("sbquo", "\u{201A}"),
    ("bdquo", "\u{201E}"),
];

/// The five characters `HtmlEncode` names, in .NET's spellings.
/// ⛔ `'` is `&#39;`, NOT `&apos;` — measured.
const SPECIALS: &[(i32, &str)] = &[
    (60, "&lt;"),
    (62, "&gt;"),
    (38, "&amp;"),
    (34, "&quot;"),
    (39, "&#39;"),
];

fn push_str(chunk: &mut Chunk, value: &str, line: u32) {
    chunk.emit_string_const(&Arc::from(value), line);
}

fn push_num(chunk: &mut Chunk, value: i32, line: u32) {
    chunk.emit_i32_const(value, line);
}

fn call(chunk: &mut Chunk, module: &str, func: &str, argc: u8, line: u32) {
    let idx = chunk.add_import(module, func);
    chunk.emit_call(idx, argc, line);
}

fn get(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn set(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}

/// `slot > value` as an `if` condition.
fn if_slot_gt(chunk: &mut Chunk, slot: u16, value: i32, line: u32) {
    get(chunk, slot, line);
    push_num(chunk, value, line);
    ops::emit_dyn_gt(chunk, line);
    chunk.emit_if(line);
}

/// `slot < value` as an `if` condition.
fn if_slot_lt(chunk: &mut Chunk, slot: u16, value: i32, line: u32) {
    get(chunk, slot, line);
    push_num(chunk, value, line);
    ops::emit_dyn_lt(chunk, line);
    chunk.emit_if(line);
}

/// `slot == value` as an `if` condition.
fn if_slot_eq_num(chunk: &mut Chunk, slot: u16, value: i32, line: u32) {
    get(chunk, slot, line);
    push_num(chunk, value, line);
    ops::emit_dyn_eq(chunk, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
}

/// `slot == value` as a VALUE-producing `if` — both arms leave one value.
///
/// ⛔ Not [`if_slot_eq_num`]: that opens a VOID block, so a value pushed inside
/// it cannot survive to the `set` after `end`.
fn if_value_slot_eq_num(chunk: &mut Chunk, slot: u16, value: i32, line: u32) {
    get(chunk, slot, line);
    push_num(chunk, value, line);
    ops::emit_dyn_eq(chunk, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
}

/// `slot == text` as an `if` condition.
fn if_slot_eq_str(chunk: &mut Chunk, slot: u16, text: &str, line: u32) {
    get(chunk, slot, line);
    push_str(chunk, text, line);
    ops::emit_dyn_eq(chunk, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
}

/// Open a `while (cursor < limit)` loop. Returns the two patch handles to pass
/// back to [`close_scan_loop`] once the body is emitted.
///
/// Block depths inside the body are loop = 0, block = 1, guard = 2, so the
/// exit is `br_if 1` and the back-edge is `br 0`.
fn open_scan_loop(chunk: &mut Chunk, cursor: u16, limit: u16, line: u32) -> (usize, usize, usize) {
    let guard = chunk.emit_block(line);
    let block = chunk.emit_block(line);
    let (loop_patch, _) = chunk.emit_loop_s(line);
    get(chunk, cursor, line);
    get(chunk, limit, line);
    ops::emit_dyn_lt(chunk, line);
    ops::emit_dyn_not(chunk, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_br_if(1, line);
    (guard, block, loop_patch)
}

fn close_scan_loop(chunk: &mut Chunk, handles: (usize, usize, usize), line: u32) {
    let (guard, block, loop_patch) = handles;
    chunk.emit_br(0, line);
    chunk.emit_end(line);
    chunk.patch_loop(loop_patch);
    chunk.emit_end(line);
    chunk.patch_block(block);
    chunk.emit_end(line);
    chunk.patch_block(guard);
}

/// `out = concat(out, piece); cursor += step`.
fn append_and_advance(chunk: &mut Chunk, out: u16, piece: u16, cursor: u16, step: u16, line: u32) {
    get(chunk, out, line);
    get(chunk, piece, line);
    call(chunk, "ecma:string", "concat", 2, line);
    set(chunk, out, line);
    get(chunk, cursor, line);
    get(chunk, step, line);
    chunk.emit_op(Op::F64_ADD, line);
    set(chunk, cursor, line);
}

/// `"&#" + String(slot) + ";"` on the stack.
fn push_numeric_reference(chunk: &mut Chunk, slot: u16, line: u32) {
    push_str(chunk, "&#", line);
    get(chunk, slot, line);
    call(chunk, "ecma:string", "String", 1, line);
    call(chunk, "ecma:string", "concat", 2, line);
    push_str(chunk, ";", line);
    call(chunk, "ecma:string", "concat", 2, line);
}

fn push_html_encode_chunk(chunks: &mut Vec<Chunk>, line: u32) -> usize {
    const NAME: &str = "__dotnet_html_encode";
    if let Some(idx) = chunks.iter().position(|chunk| chunk.name == NAME) {
        return idx;
    }
    let mut c = Chunk::new(NAME);
    c.arity = 1;
    c.local_count = 1;
    let out = c.alloc_scratch(6);
    let i = out + 1;
    let len = out + 2;
    let cp = out + 3;
    let step = out + 4;
    let piece = out + 5;

    push_str(&mut c, "", line);
    set(&mut c, out, line);
    push_num(&mut c, 0, line);
    set(&mut c, i, line);
    get(&mut c, 0, line);
    call(&mut c, "wasm:js-string", "length", 1, line);
    set(&mut c, len, line);

    let handles = open_scan_loop(&mut c, i, len, line);

    // ⛔ `codePointAt`, not `charCodeAt`: an astral character is a SURROGATE
    // PAIR in UTF-16, and .NET references the combined code point
    // (`&#128512;`). Reading code units would emit two lone-surrogate
    // references instead.
    get(&mut c, 0, line);
    get(&mut c, i, line);
    call(&mut c, "wasm:js-string", "codePointAt", 2, line);
    set(&mut c, cp, line);

    push_num(&mut c, 1, line);
    set(&mut c, step, line);
    if_slot_gt(&mut c, cp, 65535, line);
    push_num(&mut c, 2, line);
    set(&mut c, step, line);
    c.emit_end(line);

    // Default: the character itself.
    get(&mut c, cp, line);
    call(&mut c, "wasm:js-string", "fromCodePoint", 1, line);
    set(&mut c, piece, line);

    for (code, replacement) in SPECIALS {
        if_slot_eq_num(&mut c, cp, *code, line);
        push_str(&mut c, replacement, line);
        set(&mut c, piece, line);
        c.emit_end(line);
    }

    // `[160, 255]` — the Latin-1 supplement — and every astral code point.
    if_slot_gt(&mut c, cp, 159, line);
    if_slot_lt(&mut c, cp, 256, line);
    push_numeric_reference(&mut c, cp, line);
    set(&mut c, piece, line);
    c.emit_end(line);
    c.emit_end(line);
    if_slot_gt(&mut c, cp, 65535, line);
    push_numeric_reference(&mut c, cp, line);
    set(&mut c, piece, line);
    c.emit_end(line);

    append_and_advance(&mut c, out, piece, i, step, line);
    close_scan_loop(&mut c, handles, line);

    get(&mut c, out, line);
    c.emit_op(Op::RETURN, line);
    chunks.push(c);
    chunks.len() - 1
}

fn push_html_decode_chunk(chunks: &mut Vec<Chunk>, line: u32) -> usize {
    const NAME: &str = "__dotnet_html_decode";
    if let Some(idx) = chunks.iter().position(|chunk| chunk.name == NAME) {
        return idx;
    }
    let mut c = Chunk::new(NAME);
    c.arity = 1;
    c.local_count = 1;
    let out = c.alloc_scratch(11);
    let i = out + 1;
    let len = out + 2;
    let step = out + 3;
    let piece = out + 4;
    let semi = out + 5;
    let body = out + 6;
    let lowered = out + 7;
    let code = out + 8;
    let body_len = out + 9;
    let span = out + 10;

    push_str(&mut c, "", line);
    set(&mut c, out, line);
    push_num(&mut c, 0, line);
    set(&mut c, i, line);
    get(&mut c, 0, line);
    call(&mut c, "wasm:js-string", "length", 1, line);
    set(&mut c, len, line);

    let handles = open_scan_loop(&mut c, i, len, line);

    push_num(&mut c, 1, line);
    set(&mut c, step, line);
    get(&mut c, 0, line);
    get(&mut c, i, line);
    call(&mut c, "ecma:string", "charAt", 2, line);
    set(&mut c, piece, line);

    get(&mut c, 0, line);
    get(&mut c, i, line);
    call(&mut c, "wasm:js-string", "charCodeAt", 2, line);
    set(&mut c, code, line);

    if_slot_eq_num(&mut c, code, 38, line);
    // The terminating `;`. Searching from `i + 1` keeps a bare `&` from
    // matching a `;` that belongs to an earlier entity.
    get(&mut c, 0, line);
    push_str(&mut c, ";", line);
    get(&mut c, i, line);
    push_num(&mut c, 1, line);
    c.emit_op(Op::F64_ADD, line);
    call(&mut c, "ecma:string", "indexOf", 3, line);
    set(&mut c, semi, line);

    if_slot_gt(&mut c, semi, 0, line);
    // A run of more than 32 characters without a `;` is not an entity; bail
    // rather than swallow half a document into one `piece`.
    get(&mut c, semi, line);
    get(&mut c, i, line);
    c.emit_op(Op::F64_SUB, line);
    set(&mut c, span, line);
    if_slot_lt(&mut c, span, 32, line);

    get(&mut c, 0, line);
    get(&mut c, i, line);
    push_num(&mut c, 1, line);
    c.emit_op(Op::F64_ADD, line);
    get(&mut c, semi, line);
    call(&mut c, "ecma:string", "substring", 3, line);
    set(&mut c, body, line);
    get(&mut c, body, line);
    call(&mut c, "ecma:string", "toLowerCase", 1, line);
    set(&mut c, lowered, line);
    get(&mut c, lowered, line);
    call(&mut c, "wasm:js-string", "length", 1, line);
    set(&mut c, body_len, line);

    // ⛔ `charCodeAt` past the end THROWS `RangeError: index out of bounds`, it
    // does NOT answer `NaN` the way JS does. `"&;"` produces an EMPTY entity
    // body and crashed the whole program here. Every read below is guarded by
    // the length that admits it.
    if_slot_gt(&mut c, body_len, 0, line);

    get(&mut c, lowered, line);
    push_num(&mut c, 0, line);
    call(&mut c, "wasm:js-string", "charCodeAt", 2, line);
    set(&mut c, code, line);

    if_slot_eq_num(&mut c, code, 35, line);
    // `&#233;` and `&#xE9;` — `parseInt` reads both once told the radix, and
    // answers `NaN` on anything else, which leaves the text untouched.
    // `"&#;"` is body `"#"`, so the radix probe needs a second character.
    if_slot_gt(&mut c, body_len, 1, line);
    get(&mut c, lowered, line);
    push_num(&mut c, 1, line);
    call(&mut c, "wasm:js-string", "charCodeAt", 2, line);
    set(&mut c, code, line);

    if_value_slot_eq_num(&mut c, code, 120, line);
    get(&mut c, lowered, line);
    push_num(&mut c, 2, line);
    get(&mut c, body_len, line);
    call(&mut c, "ecma:string", "substring", 3, line);
    push_num(&mut c, 16, line);
    call(&mut c, "ecma:number", "parseInt", 2, line);
    c.emit_else(line);
    get(&mut c, lowered, line);
    push_num(&mut c, 1, line);
    get(&mut c, body_len, line);
    call(&mut c, "ecma:string", "substring", 3, line);
    push_num(&mut c, 10, line);
    call(&mut c, "ecma:number", "parseInt", 2, line);
    c.emit_end(line);
    set(&mut c, code, line);

    get(&mut c, code, line);
    call(&mut c, "ecma:number", "isNaN", 1, line);
    ops::emit_dyn_not(&mut c, line);
    ops::emit_dyn_to_bool(&mut c, line);
    c.emit_if(line);
    get(&mut c, code, line);
    call(&mut c, "wasm:js-string", "fromCodePoint", 1, line);
    set(&mut c, piece, line);
    get(&mut c, body_len, line);
    push_num(&mut c, 2, line);
    c.emit_op(Op::F64_ADD, line);
    set(&mut c, step, line);
    c.emit_end(line);
    c.emit_end(line); // body_len > 1

    c.emit_else(line);
    for (name, replacement) in NAMED_ENTITIES {
        if_slot_eq_str(&mut c, lowered, name, line);
        push_str(&mut c, replacement, line);
        set(&mut c, piece, line);
        get(&mut c, semi, line);
        get(&mut c, i, line);
        c.emit_op(Op::F64_SUB, line);
        push_num(&mut c, 1, line);
        c.emit_op(Op::F64_ADD, line);
        set(&mut c, step, line);
        c.emit_end(line);
    }
    c.emit_end(line);

    c.emit_end(line); // body_len > 0
    c.emit_end(line); // entity span < 32
    c.emit_end(line); // semi > 0
    c.emit_end(line); // code == '&'

    append_and_advance(&mut c, out, piece, i, step, line);
    close_scan_loop(&mut c, handles, line);

    get(&mut c, out, line);
    c.emit_op(Op::RETURN, line);
    chunks.push(c);
    chunks.len() - 1
}

/// Call a one-argument chunk. Stack: `[arg]` → `[result]`.
fn call_chunk1(chunks: &mut [Chunk], current: usize, idx: usize, line: u32) {
    let chunk = &mut chunks[current];
    let arg_slot = chunk.alloc_scratch(1);
    set(chunk, arg_slot, line);
    chunk.emit_op_u16(Op::REF_FUNC, idx as u16, line);
    chunk.emit(0, line);
    get(chunk, arg_slot, line);
    chunk.emit_op_u8_u8(Op::CALL_REF, 1, 1, line);
}

fn drop_extra_args(chunk: &mut Chunk, argc: u8, line: u32) {
    for _ in 1..argc {
        chunk.emit_op(Op::DROP, line);
    }
}

pub fn emit_html_encode(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    drop_extra_args(&mut chunks[current], argc, line);
    let idx = push_html_encode_chunk(chunks, line);
    call_chunk1(chunks, current, idx, line);
}

pub fn emit_html_decode(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    drop_extra_args(&mut chunks[current], argc, line);
    let idx = push_html_decode_chunk(chunks, line);
    call_chunk1(chunks, current, idx, line);
}

pub fn emit_url_encode(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    drop_extra_args(&mut chunks[current], argc, line);
    url::emit_percent_encode(chunks, current, URL_ENCODE_OPTIONS, line);
    // ⛔ `encodeURIComponent` leaves `'` unescaped (§19.2.6.5 lists it
    // unreserved); .NET escapes it. `PercentOptions.safe` only ever REMOVES
    // escaping, so this one is the caller's.
    let chunk = &mut chunks[current];
    push_str(chunk, "'", line);
    push_str(chunk, "%27", line);
    call(chunk, "ecma:string", "replaceAll", 3, line);
}

pub fn emit_url_decode(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    drop_extra_args(&mut chunks[current], argc, line);
    url::emit_percent_decode(chunks, current, URL_ENCODE_OPTIONS, line);
}
