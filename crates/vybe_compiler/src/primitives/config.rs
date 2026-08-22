//! `config` — INI-style configuration text as a shared primitive.
//!
//! INI is a STRUCTURED format with a real grammar — section headers, comment
//! prefixes, two accepted key/value delimiters — so it belongs beside `csv`,
//! `json` and `url` rather than inside one language's walker. It is the same
//! argument `csv.rs` makes one file over, and the same evidence: the format was
//! implemented exactly once in the tree, as a hand-written Python prelude
//! class, and no other language could reach it.
//!
//! # Who needs this
//!
//! | language | surface | before |
//! |---|---|---|
//! | python | `configparser.ConfigParser` | 75-line prelude class in `walker.rs` |
//! | php | `parse_ini_file`, `parse_ini_string` | **absent** — real PHP has both |
//!
//! Ruby (`IniFile`), .NET (`ConfigurationManager`) and Delphi (`TIniFile`) are
//! the obvious next consumers; all three spell the same three knobs.
//!
//! # The dialect is RUNTIME, not compile-time
//!
//! The same call `csv.rs` makes, for the same reason. Key CASE is where the
//! languages actually diverge: CPython's `ConfigParser.optionxform` lowercases
//! option names by default — measured, `MixedKey` reads back as `mixedkey` —
//! while PHP's `parse_ini_*` preserves them. A compile-time choice could serve
//! only one of them, so it arrives as a stack value.
//!
//! Section names are NOT folded in either language, and are not folded here.
//!
//! # Scope
//!
//! Line-oriented, which is what both consumers' surface actually is: headers,
//! `#`/`;` comments, `=`/`:` delimiters, values trimmed. Continuation lines and
//! `[DEFAULT]` inheritance are CPython-specific elaborations that no corpus
//! test exercises; they are not implemented rather than half-implemented.
//!
//! **No interpolation here.** `%(name)s` is `BasicInterpolation`, a
//! ConfigParser policy applied on READ, not part of the file format — measured
//! against CPython, `read_string` stores the raw text. Putting it here would
//! bake one language's default into the shared format.

use vybe_runtime::opcode::Op;
use vybe_runtime::Chunk;

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

/// `s == literal` as a wasm i32 condition.
fn eq_const(chunks: &mut [Chunk], current: usize, slot: u16, literal: &str, line: u32) {
    get(&mut chunks[current], slot, line);
    chunks[current].emit_string_const(literal, line);
    crate::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    crate::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
}

/// `config.parse` — INI text to a map of section name → map of key → value.
///
/// Stack: `[text, lower_keys] -> [map]`.
///
/// Keys appearing before any section header are DROPPED. CPython raises
/// `MissingSectionHeaderError` for them and PHP puts them at top level; ignoring
/// is the one behaviour neither language will mistake for its own, and the
/// alternative is inventing a third.
pub fn emit_parse(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = chunks[current].alloc_scratch(12);
    let (lower, text, out, cur, has_cur, lines, i, n, raw, t, c0, sect) = (
        base,
        base + 1,
        base + 2,
        base + 3,
        base + 4,
        base + 5,
        base + 6,
        base + 7,
        base + 8,
        base + 9,
        base + 10,
        base + 11,
    );

    // Arguments arrive in order, so they pop in reverse.
    set(&mut chunks[current], lower, line);
    set(&mut chunks[current], text, line);

    crate::primitives::collections::emit_map_new(chunks, current, line);
    set(&mut chunks[current], out, line);
    chunks[current].emit_string_const("", line);
    set(&mut chunks[current], cur, line);
    chunks[current].emit_f64_const(0.0, line);
    set(&mut chunks[current], has_cur, line);

    get(&mut chunks[current], text, line);
    chunks[current].emit_string_const("\n", line);
    call(chunks, current, "ecma:string", "split", 2, line);
    set(&mut chunks[current], lines, line);

    chunks[current].emit_f64_const(0.0, line);
    set(&mut chunks[current], i, line);
    get(&mut chunks[current], lines, line);
    chunks[current].emit_op(Op::ARRAY_LENGTH, line);
    set(&mut chunks[current], n, line);

    let loop_state = crate::primitives::loops::emit_loop_start(chunks, current, line);
    get(&mut chunks[current], i, line);
    get(&mut chunks[current], n, line);
    crate::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
    crate::primitives::loops::emit_loop_cond(chunks, current, line);

    get(&mut chunks[current], lines, line);
    get(&mut chunks[current], i, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    set(&mut chunks[current], raw, line);

    // A `\r\n` file leaves the CR on the line; `trim` takes it with the rest of
    // the surrounding space, so no separate newline dialect is needed.
    get(&mut chunks[current], raw, line);
    call(chunks, current, "ecma:string", "trim", 1, line);
    set(&mut chunks[current], t, line);

    // `charAt` past the end answers "", so this is safe on a blank line and the
    // blank test below does not have to guard it.
    get(&mut chunks[current], t, line);
    chunks[current].emit_f64_const(0.0, line);
    call(chunks, current, "ecma:string", "charAt", 2, line);
    set(&mut chunks[current], c0, line);

    // Skip blanks and comments: `if !(blank || # || ;)`.
    eq_const(chunks, current, t, "", line);
    eq_const(chunks, current, c0, "#", line);
    chunks[current].emit_op(Op::I32_OR, line);
    eq_const(chunks, current, c0, ";", line);
    chunks[current].emit_op(Op::I32_OR, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    {
        eq_const(chunks, current, c0, "[", line);
        chunks[current].emit_if(line);
        {
            // ── section header ──
            let end = chunks[current].alloc_scratch(1);
            get(&mut chunks[current], t, line);
            chunks[current].emit_string_const("]", line);
            call(chunks, current, "ecma:string", "indexOf", 2, line);
            set(&mut chunks[current], end, line);

            chunks[current].emit_f64_const(0.0, line);
            get(&mut chunks[current], end, line);
            crate::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
            crate::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
            chunks[current].emit_if(line);
            {
                get(&mut chunks[current], t, line);
                chunks[current].emit_f64_const(1.0, line);
                get(&mut chunks[current], end, line);
                call(chunks, current, "ecma:string", "slice", 3, line);
                set(&mut chunks[current], cur, line);

                // A repeated header MERGES into the existing section rather
                // than replacing it — CPython's behaviour with the default
                // `strict=False` reader, and the only one that cannot silently
                // discard keys.
                get(&mut chunks[current], out, line);
                get(&mut chunks[current], cur, line);
                call(chunks, current, "ecma:map", "has", 2, line);
                crate::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
                chunks[current].emit_op(Op::I32_EQZ, line);
                chunks[current].emit_if(line);
                {
                    get(&mut chunks[current], out, line);
                    get(&mut chunks[current], cur, line);
                    crate::primitives::collections::emit_map_new(chunks, current, line);
                    call(chunks, current, "ecma:map", "set", 3, line);
                    chunks[current].emit_op(Op::DROP, line);
                }
                chunks[current].emit_end(line);

                chunks[current].emit_f64_const(1.0, line);
                set(&mut chunks[current], has_cur, line);
            }
            chunks[current].emit_end(line);
        }
        chunks[current].emit_else(line);
        {
            // ── key/value, only once a section is open ──
            get(&mut chunks[current], has_cur, line);
            crate::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
            chunks[current].emit_if(line);
            {
                let scratch = chunks[current].alloc_scratch(4);
                let (eq_at, colon_at, at, key) =
                    (scratch, scratch + 1, scratch + 2, scratch + 3);

                get(&mut chunks[current], t, line);
                chunks[current].emit_string_const("=", line);
                call(chunks, current, "ecma:string", "indexOf", 2, line);
                set(&mut chunks[current], eq_at, line);

                get(&mut chunks[current], t, line);
                chunks[current].emit_string_const(":", line);
                call(chunks, current, "ecma:string", "indexOf", 2, line);
                set(&mut chunks[current], colon_at, line);

                // Whichever delimiter comes FIRST wins, so a value containing
                // the other one survives: `url = http://x` splits on `=`, not
                // on the `:` inside the URL.
                get(&mut chunks[current], eq_at, line);
                set(&mut chunks[current], at, line);

                get(&mut chunks[current], colon_at, line);
                chunks[current].emit_f64_const(0.0, line);
                crate::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
                crate::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
                chunks[current].emit_op(Op::I32_EQZ, line);
                chunks[current].emit_if(line);
                {
                    get(&mut chunks[current], at, line);
                    chunks[current].emit_f64_const(0.0, line);
                    crate::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
                    crate::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
                    get(&mut chunks[current], colon_at, line);
                    get(&mut chunks[current], at, line);
                    crate::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
                    crate::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
                    chunks[current].emit_op(Op::I32_OR, line);
                    chunks[current].emit_if(line);
                    {
                        get(&mut chunks[current], colon_at, line);
                        set(&mut chunks[current], at, line);
                    }
                    chunks[current].emit_end(line);
                }
                chunks[current].emit_end(line);

                // A line with neither delimiter is not a binding; drop it.
                get(&mut chunks[current], at, line);
                chunks[current].emit_f64_const(0.0, line);
                crate::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
                crate::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
                chunks[current].emit_op(Op::I32_EQZ, line);
                chunks[current].emit_if(line);
                {
                    get(&mut chunks[current], t, line);
                    chunks[current].emit_f64_const(0.0, line);
                    get(&mut chunks[current], at, line);
                    call(chunks, current, "ecma:string", "slice", 3, line);
                    call(chunks, current, "ecma:string", "trim", 1, line);
                    set(&mut chunks[current], key, line);

                    get(&mut chunks[current], lower, line);
                    crate::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
                    chunks[current].emit_if(line);
                    {
                        get(&mut chunks[current], key, line);
                        call(chunks, current, "ecma:string", "toLowerCase", 1, line);
                        set(&mut chunks[current], key, line);
                    }
                    chunks[current].emit_end(line);

                    get(&mut chunks[current], out, line);
                    get(&mut chunks[current], cur, line);
                    call(chunks, current, "ecma:map", "get", 2, line);
                    set(&mut chunks[current], sect, line);

                    get(&mut chunks[current], sect, line);
                    get(&mut chunks[current], key, line);
                    // Value runs to end of line; `slice` with one bound does
                    // that, and `trim` drops the space around the delimiter.
                    get(&mut chunks[current], t, line);
                    get(&mut chunks[current], at, line);
                    chunks[current].emit_f64_const(1.0, line);
                    chunks[current].emit_op(Op::F64_ADD, line);
                    call(chunks, current, "ecma:string", "slice", 2, line);
                    call(chunks, current, "ecma:string", "trim", 1, line);
                    call(chunks, current, "ecma:map", "set", 3, line);
                    chunks[current].emit_op(Op::DROP, line);
                }
                chunks[current].emit_end(line);
            }
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
}
