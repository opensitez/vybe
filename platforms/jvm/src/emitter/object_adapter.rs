//! `java.lang.Object.toString` — the rendering `println`, `String.valueOf` and
//! string concatenation all go through.
//!
//! `System.out.println(Object)` is specified as `String.valueOf(x)`, which is
//! `x.toString()` for a non-null reference (JLS §4.3.2 / `java.io.PrintStream`).
//! That is a JDK behaviour, so it is emitted here and every JVM frontend gets
//! the same one.
//!
//! **This replaces a syntactic rewrite.** Java used to reach a user `toString`
//! by pattern-matching the ARGUMENT EXPRESSION in the walker: a pass collected
//! the classes that declare `toString`, tracked local variable types, and — if
//! it could prove the receiver's type — renamed `x.toString()` to the canonical
//! spelling. Where the inference succeeded the rename mis-resolved, and where
//! it failed the call already worked, so the pass turned working code into
//! `[object C1]`:
//!
//! | `toString(){ … }` | before | real java |
//! |---|---|---|
//! | `return "lit";` | `lit` | `lit` |
//! | `return this.s;` | `[object C1]` | `hello` |
//! | same, receiver typed `Object` (inference fails) | `hello` | `hello` |
//!
//! A renderer has to dispatch on the VALUE, not on the expression that
//! produced it — Kotlin reached the same conclusion when it deleted
//! `wrap_printable_arg` (`languages/kotlin/src/emitter/tostring.rs`).
//!
//! The class's own `toString` is reached by its SLOT
//! ([`ProtocolSlot::ToString`]), never by spelling — flexclassplan's
//! bind-don't-name rule. Java already declares the binding
//! (`languages/java/src/protocol.rs`); nothing read it.

use vybe_compiler::primitives::{expressions, instructions::host, ops};
use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

/// `String.valueOf(x)` / `x.toString()`. Stack: `[value]` → `[string]`.
///
/// Only a non-array OBJECT can fill the slot, and `STRUCT_GET` traps on a
/// primitive, so the probe is guarded. Everything else — `null`, numbers,
/// booleans, strings, arrays — keeps the ECMA coercion it already had.
pub fn emit_to_string(chunks: &mut [Chunk], current: usize, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    host::emit(&mut chunks[current], "ecma:value", "typeof", 1, line);
    chunks[current].emit_string_const("object", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);

    // An array is an object too, and Java does NOT render one through
    // `toString` — `int[]` prints `[I@1b6d3586`. Leave it on the coercion.
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    host::emit(&mut chunks[current], "ecma:array", "isArray", 1, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    emit_ecma_string(chunks, current, value, line);
    chunks[current].emit_else(line);
    expressions::emit_rich_to_string(&mut chunks[current], value, line);
    chunks[current].emit_end(line);

    chunks[current].emit_else(line);
    emit_ecma_string(chunks, current, value, line);
    chunks[current].emit_end(line);
}

fn emit_ecma_string(chunks: &mut [Chunk], current: usize, value: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    host::emit(&mut chunks[current], "ecma:string", "String", 1, line);
}
