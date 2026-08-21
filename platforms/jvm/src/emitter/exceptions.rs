//! Shared JDK exception-throw emission for the jvm platform adapters.
//!
//! One body for every adapter that throws a named JDK exception
//! (`IllegalArgumentException`, `UnsupportedOperationException`, …). The
//! exception object is built through the shared error primitives so the name
//! is attached and a `catch (XxxException e)` matches it by class name.

use vybe_compiler::primitives::errors;
use vybe_runtime::Chunk;

/// Throw `new <name>("")`. Emits an exception construct + throw; control flow
/// does not continue past the emitted sequence at runtime.
pub(crate) fn emit_jvm_exception_throw(
    chunks: &mut [Chunk],
    current: usize,
    name: &str,
    line: u32,
) {
    emit_jvm_exception_throw_msg(chunks, current, name, "", line);
}

/// Throw `new <name>(<message>)`.
pub(crate) fn emit_jvm_exception_throw_msg(
    chunks: &mut [Chunk],
    current: usize,
    name: &str,
    message: &str,
    line: u32,
) {
    chunks[current].emit_struct_new(0, 0, line);
    chunks[current].emit_dup(line);
    chunks[current].emit_string_const(message, line);
    errors::emit_exception_new_finalize(&mut chunks[current], name, line);
    errors::emit_throw(&mut chunks[current], line);
}
