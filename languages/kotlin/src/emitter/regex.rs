//! Kotlin `kotlin.text.Regex`.
//!
//! ECMA-262 is the engine — `ecma:regexp.*` does every match — and this module
//! is the adaptation between Kotlin's surface and it.
//!
//! ## Why Kotlin does not simply reuse the JVM adapter
//!
//! It used to. `platforms/jvm`'s `regex_adapter` models `java.util.regex`, and
//! Kotlin's `Regex` is a *wrapper over* `Pattern`, so borrowing it looked free.
//! It is not, and the failures were exactly at the points where Kotlin
//! deliberately differs from the class it wraps:
//!
//! - **`split` is Kotlin's own**, not `Pattern.split`. Java drops trailing
//!   empty strings when the limit is zero; Kotlin keeps them, treats `0` as
//!   "no limit", and rejects a negative limit outright.
//! - **`RegexOption`s are a `Set`**, not an int, and the ctor takes either one
//!   option or a set of them.
//! - **`replace` takes a lambda** (`(MatchResult) -> CharSequence`), which
//!   `Pattern` has no equivalent of at all.
//! - **`MatchResult` exposes `groups`**, a collection addressable by index
//!   *and* by name; `Matcher` exposes `group(i)` calls.
//!
//! So Kotlin owns its regex ops here, reached as `common:kotlin.regex_*`, and
//! `platforms/jvm` keeps modelling Java. Java is unaffected by construction
//! rather than by measurement: nothing in this file is reachable from it.
//!
//! ## What is deliberately shared
//!
//! The compiled object's FIELD LAYOUT (`__jvm_regex_re`, `__jvm_regex_source`,
//! `__jvm_regex_flags`) is the JVM adapter's, because Kotlin's `toPattern()`
//! answers the same object — `Regex` and `Pattern` are one value here, which
//! is why `toPattern()` is a no-op — and `pattern.toPattern().matcher(s)` then
//! walks into the JVM matcher. Change the layout and that stops working.

use vybe_compiler::primitives::{collections, instructions::host, ops};
use vybe_runtime::opcode::Op;
use vybe_runtime::opcode::heaptype;
use vybe_runtime::{Chunk, Value};

use std::sync::Arc;

/// The compiled `ecma:regexp` object.
const RE_KEY: &str = "__jvm_regex_re";
/// The pattern as WRITTEN — `Regex.pattern` answers this, not the source that
/// was handed to the engine, which may have had `\Q…\E` expanded out of it.
const SOURCE_KEY: &str = "__jvm_regex_source";
/// Java's numeric flag bitmask, kept so `Pattern.flags()` can answer.
const FLAGS_KEY: &str = "__jvm_regex_flags";

// Kotlin's `RegexOption` values ARE `java.util.regex.Pattern`'s constants —
// `RegexOption.IGNORE_CASE(Pattern.CASE_INSENSITIVE)` and so on — which is why
// these are the numbers `tree_register` mounts.
const CASE_INSENSITIVE: i32 = 2;
const MULTILINE: i32 = 8;
const LITERAL: i32 = 16;
const DOTALL: i32 = 32;
const COMMENTS: i32 = 4;

fn key(chunk: &mut Chunk, name: &str) -> u16 {
    chunk.add_constant(Value::String(Arc::from(name)))
}

fn get(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn set(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}

fn field_get(chunk: &mut Chunk, slot: u16, field: &str, line: u32) {
    get(chunk, slot, line);
    let k = key(chunk, field);
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, k, line);
}

/// Stack: [value] → []. Writes `slot.field = value`.
fn field_set_from_stack(chunk: &mut Chunk, slot: u16, field: &str, line: u32) {
    let value = chunk.alloc_scratch(1);
    set(chunk, value, line);
    get(chunk, slot, line);
    get(chunk, value, line);
    let k = key(chunk, field);
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, k, line);
}

fn null(chunk: &mut Chunk, line: u32) {
    chunk.emit_ref_null(heaptype::HT_EXTERN, line);
}

/// Throw a JVM-named exception carrying `message`.
///
/// Kotlin's regex surface raises real exception TYPES — `find` past the end of
/// the input is an `IndexOutOfBoundsException`, a negative split limit is an
/// `IllegalArgumentException` — and a program catching `RuntimeException`
/// expects to see them.
fn emit_throw(chunks: &mut Vec<Chunk>, current: usize, class: &str, message: &str, line: u32) {
    chunks[current].emit_string_const(message, line);
    crate::emitter::nullability::emit_exception(chunks, current, 1, class, line);
    vybe_compiler::primitives::errors::emit_throw(&mut chunks[current], line);
}

/// `bits & mask != 0`, as an `if` condition. Stack: [] → [i32].
fn emit_flag_test(chunk: &mut Chunk, bits: u16, mask: i32, line: u32) {
    get(chunk, bits, line);
    chunk.emit_op(Op::I32_FROM_F64, line);
    chunk.emit_i32_const(mask, line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_i32_const(0, line);
    chunk.emit_op(Op::I32_NE, line);
}

/// Fold a `RegexOption` argument to Java's numeric bitmask. Stack: [] → [f64].
///
/// Kotlin overloads the constructor: `Regex(p, RegexOption.IGNORE_CASE)` passes
/// ONE option, `Regex(p, setOf(a, b))` passes a set of them. Both arrive here as
/// one value, so the shape has to be asked at runtime — `ecma:value.typeof` is
/// the spec's own answer to that question, and a `Set` (which Kotlin models as a
/// dict) reports `"object"` while a bare option is a `"number"`.
///
/// The set is drained with `iter_for_of`, the same materialisation a `for (x in
/// set)` uses, so a `Set`, a `List` or anything else iterable all work. OR
/// rather than sum: the values are distinct bits, but a caller may legitimately
/// repeat one, and `IGNORE_CASE` twice must not become `LITERAL`.
fn emit_options_bits(chunks: &mut Vec<Chunk>, current: usize, opts: u16, line: u32) {
    get(&mut chunks[current], opts, line);
    host::emit(&mut chunks[current], "ecma:value", "typeof", 1, line);
    chunks[current].emit_string_const("number", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_if_value(line);

    get(&mut chunks[current], opts, line);

    chunks[current].emit_else(line);

    get(&mut chunks[current], opts, line);
    collections::emit_iter_for_of(chunks, current, line);
    let items = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], items, line);

    get(&mut chunks[current], items, line);
    host::emit(&mut chunks[current], "ecma:array", "length", 1, line);
    chunks[current].emit_op(Op::F64_FROM_I32, line);
    let count = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], count, line);

    let acc = chunks[current].alloc_scratch(1);
    chunks[current].emit_i32_const(0, line);
    set(&mut chunks[current], acc, line);
    let index = chunks[current].alloc_scratch(1);
    chunks[current].emit_f64_const(0.0, line);
    set(&mut chunks[current], index, line);

    let block = chunks[current].emit_block(line);
    let (loop_id, _) = chunks[current].emit_loop_s(line);
    get(&mut chunks[current], index, line);
    get(&mut chunks[current], count, line);
    ops::emit_dyn_lt(&mut chunks[current], line);
    ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);

    get(&mut chunks[current], acc, line);
    get(&mut chunks[current], items, line);
    get(&mut chunks[current], index, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    chunks[current].emit_op(Op::I32_FROM_F64, line);
    chunks[current].emit_op(Op::I32_OR, line);
    set(&mut chunks[current], acc, line);

    get(&mut chunks[current], index, line);
    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    set(&mut chunks[current], index, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_id);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);

    get(&mut chunks[current], acc, line);
    chunks[current].emit_op(Op::F64_FROM_I32, line);

    chunks[current].emit_end(line);
}

/// Java's flag bitmask → an ECMA flags string. Stack: [] → [string].
///
/// Only the three flags that ECMA-262 also has are translated. `UNIX_LINES`,
/// `COMMENTS` and `CANON_EQ` have no ECMA equivalent and are NOT silently
/// dropped into a lie — they are simply absent from the engine, the same way
/// they are absent from every JS regexp, and `Pattern.flags()` still answers
/// them because the bitmask is stored whole. `LITERAL` is not a flag at all
/// here: it changes the SOURCE (see `emit_regex_new`).
fn emit_ecma_flags(chunk: &mut Chunk, bits: u16, line: u32) {
    // `d` unconditionally: `hasIndices` is what makes `exec` report each
    // group's start and end, and `MatchGroup.range` has no other source. It is
    // not observable to a Kotlin program — nothing here exposes the ECMA flags
    // string — so it costs a little work per match and buys a correct range.
    chunk.emit_string_const("d", line);
    for (mask, letter) in [
        (CASE_INSENSITIVE, "i"),
        (MULTILINE, "m"),
        (DOTALL, "s"),
    ] {
        emit_flag_test(chunk, bits, mask, line);
        chunk.emit_if_value(line);
        chunk.emit_string_const(letter, line);
        chunk.emit_else(line);
        chunk.emit_string_const("", line);
        chunk.emit_end(line);
        ops::emit_dyn_add(chunk, line);
    }
}

/// The regex metacharacters, as an ECMA character class.
///
/// Written as a source string rather than built: it is the argument to
/// `ecma:regexp.replaceAll`, and every backslash here survives into the
/// pattern the engine compiles.
const META_CLASS: &str = r"[.*+?^${}()|\[\]\\/-]";

/// Escape a runtime string so the engine matches it literally.
/// Stack: [string] → [string].
///
/// One `replaceAll` with `$&`, not a character walk: `$&` is the whole match,
/// so prefixing it with a backslash escapes whatever the class caught without
/// this code having to know which character it was.
fn emit_escape_literal(chunk: &mut Chunk, line: u32) {
    chunk.emit_string_const(META_CLASS, line);
    chunk.emit_string_const("\\$&", line);
    host::emit(chunk, "ecma:regexp", "replaceAll", 3, line);
}

/// Expand `\Q…\E` spans into escaped literals. Stack: [string] → [string].
///
/// `\Q…\E` is `java.util.regex` syntax and is not ECMA-262 syntax — the engine
/// would read `\Q` as an escaped `Q`. It has to be translated here because
/// `Regex.escape(s)` returns exactly that spelling (it IS `Pattern.quote`), and
/// `Regex(Regex.escape(s))` is the ordinary way to match a string literally.
///
/// Split on `\Q`, then each part after the first is `literal` + optional `\E` +
/// `regex`. An unterminated `\Q` quotes to the end of the pattern, which is
/// Java's rule too.
fn emit_expand_quoted(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let source = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], source, line);

    // Fast path: no `\Q` at all is the overwhelmingly common case, and it must
    // not pay for a split and a loop.
    get(&mut chunks[current], source, line);
    chunks[current].emit_string_const("\\Q", line);
    host::emit(&mut chunks[current], "ecma:string", "indexOf", 2, line);
    chunks[current].emit_op(Op::F64_FROM_I32, line);
    chunks[current].emit_f64_const(0.0, line);
    ops::emit_dyn_lt(&mut chunks[current], line);
    chunks[current].emit_if_value(line);

    get(&mut chunks[current], source, line);

    chunks[current].emit_else(line);

    get(&mut chunks[current], source, line);
    chunks[current].emit_string_const("\\Q", line);
    host::emit(&mut chunks[current], "ecma:string", "split", 2, line);
    let parts = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], parts, line);

    get(&mut chunks[current], parts, line);
    host::emit(&mut chunks[current], "ecma:array", "length", 1, line);
    chunks[current].emit_op(Op::F64_FROM_I32, line);
    let count = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], count, line);

    let out = chunks[current].alloc_scratch(1);
    get(&mut chunks[current], parts, line);
    chunks[current].emit_f64_const(0.0, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    set(&mut chunks[current], out, line);

    let index = chunks[current].alloc_scratch(1);
    chunks[current].emit_f64_const(1.0, line);
    set(&mut chunks[current], index, line);

    let block = chunks[current].emit_block(line);
    let (loop_id, _) = chunks[current].emit_loop_s(line);
    get(&mut chunks[current], index, line);
    get(&mut chunks[current], count, line);
    ops::emit_dyn_lt(&mut chunks[current], line);
    ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);

    let part = chunks[current].alloc_scratch(1);
    get(&mut chunks[current], parts, line);
    get(&mut chunks[current], index, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    set(&mut chunks[current], part, line);

    let close = chunks[current].alloc_scratch(1);
    get(&mut chunks[current], part, line);
    chunks[current].emit_string_const("\\E", line);
    host::emit(&mut chunks[current], "ecma:string", "indexOf", 2, line);
    chunks[current].emit_op(Op::F64_FROM_I32, line);
    set(&mut chunks[current], close, line);

    get(&mut chunks[current], out, line);

    get(&mut chunks[current], close, line);
    chunks[current].emit_f64_const(0.0, line);
    ops::emit_dyn_lt(&mut chunks[current], line);
    chunks[current].emit_if_value(line);

    // Unterminated `\Q`: the rest of the pattern is literal.
    get(&mut chunks[current], part, line);
    emit_escape_literal(&mut chunks[current], line);

    chunks[current].emit_else(line);

    get(&mut chunks[current], part, line);
    chunks[current].emit_f64_const(0.0, line);
    get(&mut chunks[current], close, line);
    host::emit(&mut chunks[current], "ecma:string", "slice", 3, line);
    emit_escape_literal(&mut chunks[current], line);
    get(&mut chunks[current], part, line);
    get(&mut chunks[current], close, line);
    chunks[current].emit_f64_const(2.0, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    host::emit(&mut chunks[current], "ecma:string", "slice", 2, line);
    ops::emit_dyn_add(&mut chunks[current], line);

    chunks[current].emit_end(line);

    ops::emit_dyn_add(&mut chunks[current], line);
    set(&mut chunks[current], out, line);

    get(&mut chunks[current], index, line);
    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    set(&mut chunks[current], index, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_id);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);

    get(&mut chunks[current], out, line);

    chunks[current].emit_end(line);
}

/// Strip `COMMENTS`-mode whitespace and `#` comments. Stack: [string] → [string].
///
/// Both passes lead with `(\\.)` — an escaped character — and put it back
/// through `$1`. That is what keeps `\#` and `\ ` meaningful: the alternation
/// consumes the escape before the comment or whitespace branch can see it, and
/// the replacement restores it untouched. Where the escape branch does not
/// match, `$1` expands to nothing, which is exactly the deletion wanted.
///
/// ⚠ A character class is NOT exempted. Java's `COMMENTS` ignores whitespace
/// inside `[...]` too, so that agrees; a `#` inside a class does not start a
/// comment in Java, and here it would. Recorded rather than hidden.
fn emit_strip_comments(chunk: &mut Chunk, line: u32) {
    chunk.emit_string_const(r"(\\.)|#[^\n]*", line);
    chunk.emit_string_const("$1", line);
    host::emit(chunk, "ecma:regexp", "replaceAll", 3, line);
    chunk.emit_string_const(r"(\\.)|\s+", line);
    chunk.emit_string_const("$1", line);
    host::emit(chunk, "ecma:regexp", "replaceAll", 3, line);
}

/// `Regex(pattern)` / `Regex(pattern, option)` / `Regex(pattern, options)`, and
/// `String.toRegex(…)` which is the same constructor with the receiver first.
///
/// Stack: [source] or [source, options] → [regex].
pub fn emit_regex_new(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    // Anything past the second argument is not a Kotlin overload; drop it
    // rather than leave it on the stack to corrupt the caller's frame.
    for _ in 2..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    let bits = chunks[current].alloc_scratch(1);
    if argc >= 2 {
        let opts = chunks[current].alloc_scratch(1);
        set(&mut chunks[current], opts, line);
        emit_options_bits(chunks, current, opts, line);
    } else {
        chunks[current].emit_f64_const(0.0, line);
    }
    set(&mut chunks[current], bits, line);

    let source = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], source, line);

    // `LITERAL` means the WHOLE pattern is text — `\Q…\E` inside it is text
    // too, so the two translations are alternatives, never both.
    emit_flag_test(&mut chunks[current], bits, LITERAL, line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], source, line);
    emit_escape_literal(&mut chunks[current], line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], source, line);
    emit_expand_quoted(chunks, current, line);
    chunks[current].emit_end(line);
    let compiled = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], compiled, line);

    // `COMMENTS` (Java's `/x`) is a SOURCE transformation, not a flag: the
    // pattern is re-read with whitespace and `#`-to-end-of-line comments
    // stripped. ECMA-262 has no equivalent, so it happens here, after the
    // `\Q…\E` expansion — which has already turned any literal whitespace
    // into an ESCAPED character, and escapes survive the strip.
    emit_flag_test(&mut chunks[current], bits, COMMENTS, line);
    chunks[current].emit_if(line);
    get(&mut chunks[current], compiled, line);
    emit_strip_comments(&mut chunks[current], line);
    set(&mut chunks[current], compiled, line);
    chunks[current].emit_end(line);

    // **A malformed pattern is a Kotlin exception, not a JS one.**
    //
    // `ecma:regexp.new` is right to throw a `SyntaxError` — that is what
    // ECMA-262 §22.2.3.1 says — but a Kotlin program catches
    // `PatternSyntaxException`, or one of its ancestors, and a `SyntaxError`
    // is not in that chain. Uncaught, it surfaced as the bare text `[object]`.
    //
    // So the compile is protected and the error is re-thrown in Kotlin's own
    // hierarchy. `jvm_exception_chain` carries the real Java ancestry —
    // `PatternSyntaxException` extends `IllegalArgumentException` extends
    // `RuntimeException` — so a `catch (e: RuntimeException)` matches, and a
    // `catch (e: PatternSyntaxException)` stays narrow.
    let re = chunks[current].alloc_scratch(1);
    vybe_compiler::primitives::errors::emit_try_start(&mut chunks[current], line);
    get(&mut chunks[current], compiled, line);
    emit_ecma_flags(&mut chunks[current], bits, line);
    host::emit(&mut chunks[current], "ecma:regexp", "new", 2, line);
    set(&mut chunks[current], re, line);
    vybe_compiler::primitives::errors::emit_try_end(&mut chunks[current], line);
    // The fallthrough value IS the handler block's result — the throw path
    // branches in with the exception object instead, which is how the two
    // paths are told apart below.
    null(&mut chunks[current], line);
    vybe_compiler::primitives::errors::emit_handler_block_end(&mut chunks[current], line);
    let caught = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], caught, line);
    get(&mut chunks[current], caught, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_if(line);
    get(&mut chunks[current], caught, line);
    chunks[current].emit_string_const("message", line);
    host::emit(&mut chunks[current], "ecma:object", "get", 2, line);
    crate::emitter::nullability::emit_exception(
        chunks,
        current,
        1,
        "PatternSyntaxException",
        line,
    );
    vybe_compiler::primitives::errors::emit_throw(&mut chunks[current], line);
    chunks[current].emit_end(line);

    chunks[current].emit_struct_new(0, 0, line);
    let regex = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], regex, line);
    get(&mut chunks[current], source, line);
    field_set_from_stack(&mut chunks[current], regex, SOURCE_KEY, line);
    get(&mut chunks[current], re, line);
    field_set_from_stack(&mut chunks[current], regex, RE_KEY, line);
    get(&mut chunks[current], bits, line);
    field_set_from_stack(&mut chunks[current], regex, FLAGS_KEY, line);
    get(&mut chunks[current], regex, line);
}

/// `Regex.escape(literal)` — Kotlin's is `Pattern.quote`, verbatim, and the
/// `\Q…\E` spelling is observable: a program may print it. So the string is
/// produced in Java's spelling and `emit_expand_quoted` understands it on the
/// way back in, rather than escaping character-by-character here and answering
/// something a Kotlin program would not recognise.
pub fn emit_escape(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let literal = chunk.alloc_scratch(1);
    set(chunk, literal, line);
    chunk.emit_string_const("\\Q", line);
    get(chunk, literal, line);
    ops::emit_dyn_add(chunk, line);
    chunk.emit_string_const("\\E", line);
    ops::emit_dyn_add(chunk, line);
}

/// `Regex.fromLiteral(literal)` — `Regex(literal, RegexOption.LITERAL)`.
pub fn emit_from_literal(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    chunks[current].emit_f64_const(f64::from(LITERAL), line);
    emit_regex_new(chunks, current, 2, line);
}

/// `Regex.pattern` — the source as written.
pub fn emit_pattern(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let regex = chunk.alloc_scratch(1);
    set(chunk, regex, line);
    field_get(chunk, regex, SOURCE_KEY, line);
}

/// `Regex.toPattern()` — a no-op.
///
/// Not a stub: a Kotlin `Regex` and the `java.util.regex.Pattern` it wraps are
/// ONE object here, carrying the JVM adapter's field layout, so the conversion
/// has nothing to do. `matcher()` on the result walks into that adapter.
pub fn emit_to_pattern(_chunks: &mut [Chunk], _current: usize, _line: u32) {}

// ── MatchResult ────────────────────────────────────────────────────────────
//
// Kotlin's `MatchResult` is a VALUE with properties, not a cursor with getters:
// `value`, `range`, `groupValues`, `groups` and `destructured` are all readable
// at once and stay readable after the next match. So it is built eagerly, as an
// object with those fields, and a member read is an ordinary property read —
// which is why the walker no longer rewrites `.value` into a call.

/// `start..(end-1)` as a Kotlin `IntRange`. Stack: [] → [range].
///
/// Four fields for two numbers: Kotlin spells the ends `first`/`last`, and
/// `IntRange` also implements `ClosedRange` whose ends are `start`/
/// `endInclusive`. Both spellings appear in the corpus, both are correct, and
/// answering only one would make `range.start` work and `range.first` not.
fn emit_range(chunks: &mut [Chunk], current: usize, start: u16, end: u16, line: u32) {
    let chunk = &mut chunks[current];
    chunk.emit_struct_new(0, 0, line);
    let range = chunk.alloc_scratch(1);
    set(chunk, range, line);

    get(chunk, start, line);
    field_set_from_stack(chunk, range, "first", line);
    get(chunk, start, line);
    field_set_from_stack(chunk, range, "start", line);

    get(chunk, end, line);
    chunk.emit_f64_const(1.0, line);
    chunk.emit_op(Op::F64_SUB, line);
    let last = chunk.alloc_scratch(1);
    set(chunk, last, line);
    get(chunk, last, line);
    field_set_from_stack(chunk, range, "last", line);
    get(chunk, last, line);
    field_set_from_stack(chunk, range, "endInclusive", line);

    get(chunk, range, line);
}

/// `true` when a capture did not participate in the match.
///
/// The engine answers `undefined` for those — ECMA-262 §22.2.7.2 — which is a
/// distinct value from an empty match, and the two must not be confused:
/// `(\d+)?-(\w*)` against `" -abc"` has group 1 ABSENT and group 2 EMPTY, and
/// Kotlin reports `null` for the first and `""` for the second.
fn emit_is_absent(chunk: &mut Chunk, slot: u16, line: u32) {
    get(chunk, slot, line);
    host::emit(chunk, "ecma:value", "typeof", 1, line);
    chunk.emit_string_const("undefined", line);
    ops::emit_dyn_eq(chunk, line);
}

/// A Kotlin `MatchGroup` — `{ value, range }` — or `null` when the group did
/// not participate. Stack: [] → [group | null].
///
/// `indices` is the `d` flag's per-group `[start, end]` pair. It is why the
/// compiled regexp always carries `d`: without it a group's range would have to
/// be guessed from a substring search, which is wrong the moment the same text
/// appears twice.
fn emit_group_object(
    chunks: &mut [Chunk],
    current: usize,
    value: u16,
    pair: u16,
    offset: u16,
    line: u32,
) {
    emit_is_absent(&mut chunks[current], value, line);
    chunks[current].emit_if_value(line);
    null(&mut chunks[current], line);
    chunks[current].emit_else(line);

    chunks[current].emit_struct_new(0, 0, line);
    let group = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], group, line);
    get(&mut chunks[current], value, line);
    field_set_from_stack(&mut chunks[current], group, "value", line);

    // A group can carry a value and still have no indices entry if the engine
    // was compiled without `d`; answering a zero range beats trapping.
    let start = chunks[current].alloc_scratch(1);
    let end = chunks[current].alloc_scratch(1);
    get(&mut chunks[current], pair, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_f64_const(0.0, line);
    chunks[current].emit_f64_const(0.0, line);
    set(&mut chunks[current], end, line);
    set(&mut chunks[current], start, line);
    chunks[current].emit_f64_const(0.0, line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], pair, line);
    chunks[current].emit_f64_const(0.0, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    host::emit(&mut chunks[current], "wasm:js-number", "toF64", 1, line);
    get(&mut chunks[current], offset, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    set(&mut chunks[current], start, line);
    get(&mut chunks[current], pair, line);
    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    host::emit(&mut chunks[current], "wasm:js-number", "toF64", 1, line);
    get(&mut chunks[current], offset, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    set(&mut chunks[current], end, line);
    chunks[current].emit_f64_const(0.0, line);
    chunks[current].emit_end(line);
    chunks[current].emit_op(Op::DROP, line);

    emit_range(chunks, current, start, end, line);
    field_set_from_stack(&mut chunks[current], group, "range", line);
    get(&mut chunks[current], group, line);

    chunks[current].emit_end(line);
}

/// Store `value` (on the stack) at `key` (a slot) in `dict`, keeping the
/// `__keys` sidecar the dict primitives enumerate from in step.
fn dict_put(chunks: &mut [Chunk], current: usize, dict: u16, key_slot: u16, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], value, line);
    get(&mut chunks[current], dict, line);
    get(&mut chunks[current], key_slot, line);
    get(&mut chunks[current], value, line);
    // ⛔ `ARRAY_SET` pops three and pushes NOTHING on the ordinary path
    // (`dispatch.rs`) — only the `__setitem__` dunder path leaves a value.
    // A `DROP` here therefore ate whatever the CALLER had on the stack, so
    // `"text " + regex.find(s)!!.value` printed the match with `null` in front
    // of it: the literal was the value that got dropped. Invisible to the whole
    // corpus, because every test binds the match to a variable first.
    chunks[current].emit_op(Op::ARRAY_SET, line);
    get(&mut chunks[current], dict, line);
    let keys = key(&mut chunks[current], "__keys");
    chunks[current].emit_struct_field_op(Op::STRUCT_GET, 0, keys, line);
    get(&mut chunks[current], key_slot, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
}

/// Store `value` (on the stack) at `key`, WITHOUT registering the key for
/// enumeration. Stack: [value] → [].
fn dict_put_unkeyed(chunks: &mut [Chunk], current: usize, dict: u16, key_slot: u16, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], value, line);
    get(&mut chunks[current], dict, line);
    get(&mut chunks[current], key_slot, line);
    get(&mut chunks[current], value, line);
    chunks[current].emit_op(Op::ARRAY_SET, line);
}

/// Build a `MatchResult` from an engine match. Stack: [] → [MatchResult].
///
/// `m` is the `ecma:regexp.exec` result; `offset` is where the searched
/// substring began in the ORIGINAL input, because `find(input, startIndex)`
/// searches a slice and every index Kotlin reports is against the whole string.
fn emit_match_result(chunks: &mut Vec<Chunk>, current: usize, m: u16, offset: u16, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    get(&mut chunks[current], m, line);
    chunks[current].emit_f64_const(0.0, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    set(&mut chunks[current], value, line);

    let start = chunks[current].alloc_scratch(1);
    get(&mut chunks[current], offset, line);
    get(&mut chunks[current], m, line);
    chunks[current].emit_string_const("index", line);
    host::emit(&mut chunks[current], "ecma:object", "get", 2, line);
    host::emit(&mut chunks[current], "wasm:js-number", "toF64", 1, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    set(&mut chunks[current], start, line);

    let end = chunks[current].alloc_scratch(1);
    get(&mut chunks[current], start, line);
    get(&mut chunks[current], value, line);
    host::emit(&mut chunks[current], "wasm:js-string", "length", 1, line);
    chunks[current].emit_op(Op::F64_FROM_I32, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    set(&mut chunks[current], end, line);

    chunks[current].emit_struct_new(0, 0, line);
    let result = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], result, line);
    get(&mut chunks[current], value, line);
    field_set_from_stack(&mut chunks[current], result, "value", line);
    emit_range(chunks, current, start, end, line);
    field_set_from_stack(&mut chunks[current], result, "range", line);

    let indices = chunks[current].alloc_scratch(1);
    get(&mut chunks[current], m, line);
    chunks[current].emit_string_const("indices", line);
    host::emit(&mut chunks[current], "ecma:object", "get", 2, line);
    set(&mut chunks[current], indices, line);

    let values = chunks[current].alloc_scratch(1);
    collections::emit_array_new(chunks, current, 0, line);
    set(&mut chunks[current], values, line);

    let groups = chunks[current].alloc_scratch(1);
    new_dict(chunks, current, line);
    set(&mut chunks[current], groups, line);

    let count = chunks[current].alloc_scratch(1);
    get(&mut chunks[current], m, line);
    host::emit(&mut chunks[current], "ecma:array", "length", 1, line);
    chunks[current].emit_op(Op::F64_FROM_I32, line);
    set(&mut chunks[current], count, line);

    let index = chunks[current].alloc_scratch(1);
    chunks[current].emit_f64_const(0.0, line);
    set(&mut chunks[current], index, line);

    let raw = chunks[current].alloc_scratch(1);
    let pair = chunks[current].alloc_scratch(1);

    let block = chunks[current].emit_block(line);
    let (loop_id, _) = chunks[current].emit_loop_s(line);
    get(&mut chunks[current], index, line);
    get(&mut chunks[current], count, line);
    ops::emit_dyn_lt(&mut chunks[current], line);
    ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);

    get(&mut chunks[current], m, line);
    get(&mut chunks[current], index, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    set(&mut chunks[current], raw, line);

    // `groupValues` never holds `null`: Kotlin fills a non-participating group
    // with the empty string there, and reports the absence through `groups`.
    get(&mut chunks[current], values, line);
    emit_is_absent(&mut chunks[current], raw, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_string_const("", line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], raw, line);
    chunks[current].emit_end(line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    get(&mut chunks[current], indices, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if_value(line);
    null(&mut chunks[current], line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], indices, line);
    get(&mut chunks[current], index, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    chunks[current].emit_end(line);
    set(&mut chunks[current], pair, line);

    emit_group_object(chunks, current, raw, pair, offset, line);
    dict_put(chunks, current, groups, index, line);

    get(&mut chunks[current], index, line);
    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    set(&mut chunks[current], index, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_id);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);

    // Named groups. `(?<id>…)` is addressable by NAME as well as by ordinal —
    // the same `MatchGroup`, reached two ways — so the dict carries both keys.
    let named = chunks[current].alloc_scratch(1);
    get(&mut chunks[current], m, line);
    chunks[current].emit_string_const("groups", line);
    host::emit(&mut chunks[current], "ecma:object", "get", 2, line);
    set(&mut chunks[current], named, line);

    get(&mut chunks[current], named, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_if(line);

    let named_indices = chunks[current].alloc_scratch(1);
    get(&mut chunks[current], indices, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if_value(line);
    null(&mut chunks[current], line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], indices, line);
    chunks[current].emit_string_const("groups", line);
    host::emit(&mut chunks[current], "ecma:object", "get", 2, line);
    chunks[current].emit_end(line);
    set(&mut chunks[current], named_indices, line);

    let names = chunks[current].alloc_scratch(1);
    get(&mut chunks[current], named, line);
    host::emit(&mut chunks[current], "ecma:object", "keys", 1, line);
    set(&mut chunks[current], names, line);

    let name_count = chunks[current].alloc_scratch(1);
    get(&mut chunks[current], names, line);
    host::emit(&mut chunks[current], "ecma:array", "length", 1, line);
    chunks[current].emit_op(Op::F64_FROM_I32, line);
    set(&mut chunks[current], name_count, line);

    let n = chunks[current].alloc_scratch(1);
    chunks[current].emit_f64_const(0.0, line);
    set(&mut chunks[current], n, line);
    let name = chunks[current].alloc_scratch(1);

    let nblock = chunks[current].emit_block(line);
    let (nloop, _) = chunks[current].emit_loop_s(line);
    get(&mut chunks[current], n, line);
    get(&mut chunks[current], name_count, line);
    ops::emit_dyn_lt(&mut chunks[current], line);
    ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);

    get(&mut chunks[current], names, line);
    get(&mut chunks[current], n, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    set(&mut chunks[current], name, line);

    get(&mut chunks[current], named, line);
    get(&mut chunks[current], name, line);
    host::emit(&mut chunks[current], "ecma:object", "get", 2, line);
    set(&mut chunks[current], raw, line);

    get(&mut chunks[current], named_indices, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if_value(line);
    null(&mut chunks[current], line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], named_indices, line);
    get(&mut chunks[current], name, line);
    host::emit(&mut chunks[current], "ecma:object", "get", 2, line);
    chunks[current].emit_end(line);
    set(&mut chunks[current], pair, line);

    emit_group_object(chunks, current, raw, pair, offset, line);
    // NOT `dict_put`: a name is a second way to REACH a group, not another
    // group. `MatchGroupCollection.size` is the ordinal count — 3 for two
    // capture groups — and it reads `__keys`, so registering the names there
    // would report 5 for `(?<a>…)(?<b>…)`.
    dict_put_unkeyed(chunks, current, groups, name, line);

    get(&mut chunks[current], n, line);
    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    set(&mut chunks[current], n, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(nloop);
    chunks[current].emit_end(line);
    chunks[current].patch_block(nblock);

    chunks[current].emit_end(line);

    get(&mut chunks[current], values, line);
    field_set_from_stack(&mut chunks[current], result, "groupValues", line);
    get(&mut chunks[current], groups, line);
    field_set_from_stack(&mut chunks[current], result, "groups", line);
    // `destructured` drops group 0 and is otherwise `groupValues` — built from
    // the FILLED array, so `component1()` on an absent group is `""` and not
    // `undefined`.
    get(&mut chunks[current], values, line);
    chunks[current].emit_f64_const(1.0, line);
    host::emit(&mut chunks[current], "ecma:array", "slice", 2, line);
    field_set_from_stack(&mut chunks[current], result, "destructured", line);
    get(&mut chunks[current], result, line);
}

/// An empty dict with the `__keys` sidecar the enumeration primitives read.
fn new_dict(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    chunks[current].emit_struct_new(0, 0, line);
    chunks[current].emit_dup(line);
    collections::emit_array_new(chunks, current, 0, line);
    let keys = key(&mut chunks[current], "__keys");
    chunks[current].emit_struct_field_op(Op::STRUCT_SET, 0, keys, line);
}

// ── Searching ──────────────────────────────────────────────────────────────

/// Run the engine at `from` in `input`, leaving the raw match (or null) in
/// `out`. Stack: [] → [].
fn emit_exec_at(
    chunks: &mut [Chunk],
    current: usize,
    regex: u16,
    input: u16,
    from: u16,
    out: u16,
    line: u32,
) {
    let chunk = &mut chunks[current];
    field_get(chunk, regex, RE_KEY, line);
    get(chunk, input, line);
    get(chunk, from, line);
    host::emit(chunk, "ecma:string", "slice", 2, line);
    host::emit(chunk, "ecma:regexp", "exec", 2, line);
    set(chunk, out, line);
}

/// `input.length`, as an f64. Stack: [] → [f64].
fn emit_length_of(chunk: &mut Chunk, slot: u16, line: u32) {
    get(chunk, slot, line);
    host::emit(chunk, "wasm:js-string", "length", 1, line);
    chunk.emit_op(Op::F64_FROM_I32, line);
}

/// Pop `[regex, input]` or `[regex, input, start]` into slots, defaulting the
/// start index to 0. Returns `(regex, input, start)`.
fn pop_search_args(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) -> (u16, u16, u16) {
    let chunk = &mut chunks[current];
    let start = chunk.alloc_scratch(1);
    if argc >= 3 {
        set(chunk, start, line);
    }
    let input = chunk.alloc_scratch(1);
    let regex = chunk.alloc_scratch(1);
    set(chunk, input, line);
    set(chunk, regex, line);
    if argc < 3 {
        chunk.emit_f64_const(0.0, line);
        set(chunk, start, line);
    }
    (regex, input, start)
}

/// Kotlin rejects a start index outside the input rather than clamping it:
/// `find(input, input.length)` is legal and answers null, `input.length + 1`
/// throws. Verified against the Kotlin compiler, not assumed — the corpus
/// contains a test asserting the opposite, and the oracle settled it.
fn emit_guard_start_index(chunks: &mut Vec<Chunk>, current: usize, input: u16, start: u16, line: u32) {
    emit_length_of(&mut chunks[current], input, line);
    get(&mut chunks[current], start, line);
    ops::emit_dyn_lt(&mut chunks[current], line);
    chunks[current].emit_if(line);
    emit_throw(chunks, current, "IndexOutOfBoundsException", "Illegal start index", line);
    chunks[current].emit_end(line);
}

/// `Regex.find(input, startIndex = 0)` → `MatchResult?`
pub fn emit_find(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let (regex, input, start) = pop_search_args(chunks, current, argc, line);
    emit_guard_start_index(chunks, current, input, start, line);
    let m = chunks[current].alloc_scratch(1);
    emit_exec_at(chunks, current, regex, input, start, m, line);
    get(&mut chunks[current], m, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if_value(line);
    null(&mut chunks[current], line);
    chunks[current].emit_else(line);
    emit_match_result(chunks, current, m, start, line);
    chunks[current].emit_end(line);
}

/// `Regex.findAll(input, startIndex = 0)` → every match, left to right.
///
/// Kotlin's answer is a `Sequence`, and this is a `List`. That is not a
/// shortcut taken to avoid laziness: Kotlin's sequence here is *constrained to
/// be iterated once* only for `generateSequence`-style sources, while
/// `findAll`'s is re-iterable, so `count()` followed by `joinToString` — which
/// the corpus does — works on both. The observable difference is exhaustion,
/// and this one cannot be exhausted.
pub fn emit_find_all(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let (regex, input, start) = pop_search_args(chunks, current, argc, line);
    emit_guard_start_index(chunks, current, input, start, line);

    let out = chunks[current].alloc_scratch(1);
    collections::emit_array_new(chunks, current, 0, line);
    set(&mut chunks[current], out, line);

    let cursor = chunks[current].alloc_scratch(1);
    get(&mut chunks[current], start, line);
    set(&mut chunks[current], cursor, line);

    let m = chunks[current].alloc_scratch(1);
    let limit = chunks[current].alloc_scratch(1);
    emit_length_of(&mut chunks[current], input, line);
    set(&mut chunks[current], limit, line);

    let block = chunks[current].emit_block(line);
    let (loop_id, _) = chunks[current].emit_loop_s(line);

    get(&mut chunks[current], limit, line);
    get(&mut chunks[current], cursor, line);
    ops::emit_dyn_lt(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);

    emit_exec_at(chunks, current, regex, input, cursor, m, line);
    get(&mut chunks[current], m, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_br_if(1, line);

    let result = chunks[current].alloc_scratch(1);
    emit_match_result(chunks, current, m, cursor, line);
    set(&mut chunks[current], result, line);
    get(&mut chunks[current], out, line);
    get(&mut chunks[current], result, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    // A zero-width match would otherwise re-match at the same place for ever.
    // Stepping one position past it is what every engine does, and it is why
    // `Regex("a*").findAll("bab")` reports four matches and not an infinite
    // number of them.
    let next = chunks[current].alloc_scratch(1);
    get(&mut chunks[current], result, line);
    chunks[current].emit_string_const("range", line);
    host::emit(&mut chunks[current], "ecma:object", "get", 2, line);
    chunks[current].emit_string_const("last", line);
    host::emit(&mut chunks[current], "ecma:object", "get", 2, line);
    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    set(&mut chunks[current], next, line);

    get(&mut chunks[current], next, line);
    get(&mut chunks[current], cursor, line);
    ops::emit_dyn_gt(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], next, line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], cursor, line);
    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    chunks[current].emit_end(line);
    set(&mut chunks[current], cursor, line);

    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_id);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);

    get(&mut chunks[current], out, line);
}

/// `Regex.matchEntire(input)` — a match only if it spans the WHOLE input.
pub fn emit_match_entire(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let input = chunks[current].alloc_scratch(1);
    let regex = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], input, line);
    set(&mut chunks[current], regex, line);

    let zero = chunks[current].alloc_scratch(1);
    chunks[current].emit_f64_const(0.0, line);
    set(&mut chunks[current], zero, line);

    let m = chunks[current].alloc_scratch(1);
    emit_exec_at(chunks, current, regex, input, zero, m, line);

    get(&mut chunks[current], m, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if_value(line);
    null(&mut chunks[current], line);
    chunks[current].emit_else(line);

    get(&mut chunks[current], m, line);
    chunks[current].emit_f64_const(0.0, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    get(&mut chunks[current], input, line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    ops::emit_i32_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    emit_match_result(chunks, current, m, zero, line);
    chunks[current].emit_else(line);
    null(&mut chunks[current], line);
    chunks[current].emit_end(line);

    chunks[current].emit_end(line);
}

/// `Regex.matches(input)` — the whole input, as a Boolean.
pub fn emit_matches(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    emit_match_entire(chunks, current, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    ops::emit_dyn_not(&mut chunks[current], line);
    ops::emit_i32_to_bool(&mut chunks[current], line);
}

/// `Regex.containsMatchIn(input)` — a match ANYWHERE.
pub fn emit_contains(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let input = chunk.alloc_scratch(1);
    let regex = chunk.alloc_scratch(1);
    set(chunk, input, line);
    set(chunk, regex, line);
    field_get(chunk, regex, RE_KEY, line);
    get(chunk, input, line);
    host::emit(chunk, "ecma:regexp", "test", 2, line);
}

/// `Regex.matchesAt(input, index)` — a match STARTING exactly at `index`.
pub fn emit_matches_at(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let (regex, input, start) = pop_search_args(chunks, current, argc, line);
    let m = chunks[current].alloc_scratch(1);
    emit_exec_at(chunks, current, regex, input, start, m, line);
    get(&mut chunks[current], m, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_bool_const(false, line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], m, line);
    chunks[current].emit_string_const("index", line);
    host::emit(&mut chunks[current], "ecma:object", "get", 2, line);
    host::emit(&mut chunks[current], "wasm:js-number", "toF64", 1, line);
    chunks[current].emit_f64_const(0.0, line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    ops::emit_i32_to_bool(&mut chunks[current], line);
    chunks[current].emit_end(line);
}

// ── split ──────────────────────────────────────────────────────────────────

/// `Regex.split(input, limit = 0)` — Kotlin's own, not `Pattern.split`.
///
/// Three differences from the Java method it is often mistaken for, all
/// checked against the Kotlin compiler:
///
/// - **`limit = 0` means NO limit**, not "drop trailing empties". Java's
///   `split` uses zero to mean both at once; Kotlin separated them.
/// - **Trailing empty strings are KEPT.** `Regex(",").split("a,b,c,")` is four
///   elements ending in `""`, where Java's answers three.
/// - **A negative limit is rejected**, where Java reads it as "no limit and
///   keep everything".
///
/// So it cannot go through `ecma:regexp.split` either: JS truncates at the
/// limit (`"a,b,c,d".split(/,/, 2)` is `["a","b"]`) where Kotlin collects the
/// remainder into the last element (`["a","b,c,d"]`).
pub fn emit_split(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let limit = chunks[current].alloc_scratch(1);
    if argc >= 3 {
        set(&mut chunks[current], limit, line);
    }
    let input = chunks[current].alloc_scratch(1);
    let regex = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], input, line);
    set(&mut chunks[current], regex, line);
    if argc < 3 {
        chunks[current].emit_f64_const(0.0, line);
        set(&mut chunks[current], limit, line);
    }

    chunks[current].emit_f64_const(0.0, line);
    get(&mut chunks[current], limit, line);
    ops::emit_dyn_gt(&mut chunks[current], line);
    chunks[current].emit_if(line);
    emit_throw(chunks, current, "IllegalArgumentException", "Limit must be non-negative", line);
    chunks[current].emit_end(line);

    let out = chunks[current].alloc_scratch(1);
    collections::emit_array_new(chunks, current, 0, line);
    set(&mut chunks[current], out, line);

    let last = chunks[current].alloc_scratch(1);
    chunks[current].emit_f64_const(0.0, line);
    set(&mut chunks[current], last, line);
    let cursor = chunks[current].alloc_scratch(1);
    chunks[current].emit_f64_const(0.0, line);
    set(&mut chunks[current], cursor, line);
    let taken = chunks[current].alloc_scratch(1);
    chunks[current].emit_f64_const(0.0, line);
    set(&mut chunks[current], taken, line);

    let length = chunks[current].alloc_scratch(1);
    emit_length_of(&mut chunks[current], input, line);
    set(&mut chunks[current], length, line);

    let m = chunks[current].alloc_scratch(1);
    let start = chunks[current].alloc_scratch(1);
    let end = chunks[current].alloc_scratch(1);

    let block = chunks[current].emit_block(line);
    let (loop_id, _) = chunks[current].emit_loop_s(line);

    // One short of the limit: the rest of the input is the final element, so
    // stop splitting and let the tail push below collect it.
    get(&mut chunks[current], limit, line);
    chunks[current].emit_f64_const(0.0, line);
    ops::emit_dyn_gt(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], taken, line);
    get(&mut chunks[current], limit, line);
    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_op(Op::F64_SUB, line);
    ops::emit_dyn_ge(&mut chunks[current], line);
    chunks[current].emit_else(line);
    chunks[current].emit_bool_const(false, line);
    chunks[current].emit_end(line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);

    get(&mut chunks[current], length, line);
    get(&mut chunks[current], cursor, line);
    ops::emit_dyn_lt(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);

    emit_exec_at(chunks, current, regex, input, cursor, m, line);
    get(&mut chunks[current], m, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_br_if(1, line);

    get(&mut chunks[current], cursor, line);
    get(&mut chunks[current], m, line);
    chunks[current].emit_string_const("index", line);
    host::emit(&mut chunks[current], "ecma:object", "get", 2, line);
    host::emit(&mut chunks[current], "wasm:js-number", "toF64", 1, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    set(&mut chunks[current], start, line);

    get(&mut chunks[current], start, line);
    get(&mut chunks[current], m, line);
    chunks[current].emit_f64_const(0.0, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    host::emit(&mut chunks[current], "wasm:js-string", "length", 1, line);
    chunks[current].emit_op(Op::F64_FROM_I32, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    set(&mut chunks[current], end, line);

    // ⛔ A zero-width match at the END of the input IS a separator. I guarded
    // against it first, reasoning that it would add an element Kotlin does not
    // report; kotlinc says otherwise — `Regex("").split("abc")` is FIVE
    // elements, `["", "a", "b", "c", ""]`, and `Regex("a*").split("bab")` is
    // `["", "b", "", "b", ""]`. Kotlin's split walks `findAll` and appends the
    // tail unconditionally, so the final empty match contributes a piece and
    // the tail contributes another. The loop still terminates because `cursor`
    // strictly increases: a zero-width match advances it by one.
    get(&mut chunks[current], out, line);
    get(&mut chunks[current], input, line);
    get(&mut chunks[current], last, line);
    get(&mut chunks[current], start, line);
    host::emit(&mut chunks[current], "ecma:string", "slice", 3, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    get(&mut chunks[current], taken, line);
    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    set(&mut chunks[current], taken, line);

    get(&mut chunks[current], end, line);
    set(&mut chunks[current], last, line);

    get(&mut chunks[current], end, line);
    get(&mut chunks[current], cursor, line);
    ops::emit_dyn_gt(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], end, line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], cursor, line);
    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    chunks[current].emit_end(line);
    set(&mut chunks[current], cursor, line);

    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_id);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);

    get(&mut chunks[current], out, line);
    get(&mut chunks[current], input, line);
    get(&mut chunks[current], last, line);
    host::emit(&mut chunks[current], "ecma:string", "slice", 2, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    get(&mut chunks[current], out, line);
}

// ── replace ────────────────────────────────────────────────────────────────

/// `Regex.replace(input, replacement)` and `Regex.replace(input) { … }`.
///
/// Kotlin overloads the same name on a String and on a
/// `(MatchResult) -> CharSequence`, and the two have nothing in common at
/// runtime: the string form is a template the engine expands (`$1`, `$&`),
/// while the lambda form has to be CALLED once per match with a real
/// `MatchResult`. `ecma:value.typeof` picks, for the same reason the option
/// fold uses it — it is the spec's own question about a value's shape.
pub fn emit_replace(chunks: &mut Vec<Chunk>, current: usize, first_only: bool, line: u32) {
    let with = chunks[current].alloc_scratch(1);
    let input = chunks[current].alloc_scratch(1);
    let regex = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], with, line);
    set(&mut chunks[current], input, line);
    set(&mut chunks[current], regex, line);

    get(&mut chunks[current], with, line);
    host::emit(&mut chunks[current], "ecma:value", "typeof", 1, line);
    chunks[current].emit_string_const("function", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_if_value(line);

    emit_replace_with_transform(chunks, current, regex, input, with, first_only, line);

    chunks[current].emit_else(line);

    get(&mut chunks[current], input, line);
    field_get(&mut chunks[current], regex, RE_KEY, line);
    get(&mut chunks[current], with, line);
    host::emit(
        &mut chunks[current],
        "ecma:regexp",
        if first_only { "replace" } else { "replaceAll" },
        3,
        line,
    );

    chunks[current].emit_end(line);
}

fn emit_replace_with_transform(
    chunks: &mut Vec<Chunk>,
    current: usize,
    regex: u16,
    input: u16,
    transform: u16,
    first_only: bool,
    line: u32,
) {
    let out = chunks[current].alloc_scratch(1);
    chunks[current].emit_string_const("", line);
    set(&mut chunks[current], out, line);

    let last = chunks[current].alloc_scratch(1);
    chunks[current].emit_f64_const(0.0, line);
    set(&mut chunks[current], last, line);
    let cursor = chunks[current].alloc_scratch(1);
    chunks[current].emit_f64_const(0.0, line);
    set(&mut chunks[current], cursor, line);

    let length = chunks[current].alloc_scratch(1);
    emit_length_of(&mut chunks[current], input, line);
    set(&mut chunks[current], length, line);

    let m = chunks[current].alloc_scratch(1);
    let result = chunks[current].alloc_scratch(1);
    let start = chunks[current].alloc_scratch(1);
    let end = chunks[current].alloc_scratch(1);

    let block = chunks[current].emit_block(line);
    let (loop_id, _) = chunks[current].emit_loop_s(line);

    get(&mut chunks[current], length, line);
    get(&mut chunks[current], cursor, line);
    ops::emit_dyn_lt(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);

    emit_exec_at(chunks, current, regex, input, cursor, m, line);
    get(&mut chunks[current], m, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_br_if(1, line);

    emit_match_result(chunks, current, m, cursor, line);
    set(&mut chunks[current], result, line);

    get(&mut chunks[current], result, line);
    chunks[current].emit_string_const("range", line);
    host::emit(&mut chunks[current], "ecma:object", "get", 2, line);
    let range = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], range, line);
    get(&mut chunks[current], range, line);
    chunks[current].emit_string_const("first", line);
    host::emit(&mut chunks[current], "ecma:object", "get", 2, line);
    set(&mut chunks[current], start, line);
    get(&mut chunks[current], range, line);
    chunks[current].emit_string_const("last", line);
    host::emit(&mut chunks[current], "ecma:object", "get", 2, line);
    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    set(&mut chunks[current], end, line);

    get(&mut chunks[current], out, line);
    get(&mut chunks[current], input, line);
    get(&mut chunks[current], last, line);
    get(&mut chunks[current], start, line);
    host::emit(&mut chunks[current], "ecma:string", "slice", 3, line);
    ops::emit_dyn_add(&mut chunks[current], line);

    get(&mut chunks[current], transform, line);
    get(&mut chunks[current], result, line);
    vybe_compiler::primitives::callable::emit_direct_invoke(chunks, current, 1, line);
    ops::emit_dyn_add(&mut chunks[current], line);
    set(&mut chunks[current], out, line);

    get(&mut chunks[current], end, line);
    set(&mut chunks[current], last, line);

    get(&mut chunks[current], end, line);
    get(&mut chunks[current], cursor, line);
    ops::emit_dyn_gt(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], end, line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], cursor, line);
    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    chunks[current].emit_end(line);
    set(&mut chunks[current], cursor, line);

    if first_only {
        chunks[current].emit_br(1, line);
    } else {
        chunks[current].emit_br(0, line);
    }
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_id);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);

    get(&mut chunks[current], out, line);
    get(&mut chunks[current], input, line);
    get(&mut chunks[current], last, line);
    host::emit(&mut chunks[current], "ecma:string", "slice", 2, line);
    ops::emit_dyn_add(&mut chunks[current], line);
}
