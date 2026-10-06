// Frozen dynamic-comparison body with only the five string predicate repairs.
// Performance reference for cursor changes; original buggy body remains separate.
use vybe_compiler::primitives::strings::{GlobOptions, emit_index_of, emit_length, emit_substring};
use vybe_runtime::{Chunk, Op};
fn get(c: &mut Chunk, slot: u16, line: u32) {
    c.emit_op_u16(Op::LOCAL_GET, slot, line);
}
fn set(c: &mut Chunk, slot: u16, line: u32) {
    c.emit_op_u16(Op::LOCAL_SET, slot, line);
}
fn i32c(c: &mut Chunk, value: i32, line: u32) {
    c.emit_i32_const(value, line);
}
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

    let loop_state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    get(&mut chunks[current], i, line);
    get(&mut chunks[current], n, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
    vybe_compiler::primitives::loops::emit_loop_cond_from_i32(chunks, current, line);

    // c = pat[i]
    emit_char_at(chunks, current, pat, i, line);
    set(&mut chunks[current], c, line);

    get(&mut chunks[current], in_class, line);
    chunks[current].emit_if_value(line);
    {
        // Inside `[...]` everything is literal until the closing `]`.
        get(&mut chunks[current], c, line);
        chunks[current].emit_string_const("]", line);
        let equals = chunks[current].add_import("wasm:js-string", "equals");
        chunks[current].emit_call(equals, 2, line);
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
        let equals = chunks[current].add_import("wasm:js-string", "equals");
        chunks[current].emit_call(equals, 2, line);
        chunks[current].emit_if_value(line);
        chunks[current].emit_string_const(".*", line);
        chunks[current].emit_else(line);

        get(&mut chunks[current], c, line);
        chunks[current].emit_string_const("?", line);
        let equals = chunks[current].add_import("wasm:js-string", "equals");
        chunks[current].emit_call(equals, 2, line);
        chunks[current].emit_if_value(line);
        chunks[current].emit_string_const(".", line);
        chunks[current].emit_else(line);

        get(&mut chunks[current], c, line);
        chunks[current].emit_string_const("[", line);
        let equals = chunks[current].add_import("wasm:js-string", "equals");
        chunks[current].emit_call(equals, 2, line);
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
            let equals = chunks[current].add_import("wasm:js-string", "equals");
            chunks[current].emit_call(equals, 2, line);
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
            vybe_compiler::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
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
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, loop_state, line);

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
    vybe_compiler::primitives::ops::emit_dyn_ne(&mut chunks[current], line);
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
