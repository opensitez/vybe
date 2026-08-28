//! Names a Python module exposes that Vybe can only partly implement.
//!
//! Two shapes, chosen by what the name IS in CPython:
//!
//! * `python.typeobj.<Name>` — a TYPE (`types.UnionType`, `numbers.Integral`,
//!   `cmd.Cmd`). Reading it yields a type object carrying `__name__`;
//!   constructing one raises, because the class behind it is not implemented.
//! * `python.surface.<name>` — a FUNCTION (`dbm.open`, `tarfile.open`,
//!   `doctest.testmod`). Reading it yields a callable; CALLING it raises
//!   `NotImplementedError` naming the module, rather than returning something
//!   that looks like it worked.
//!
//! The argument count is known when this emits, so "read" and "call" are
//! distinguished statically: `argc == 0` is the reference, anything else is a
//! call. Nothing here silently succeeds — a use that cannot be honoured raises
//! at the call site, loudly, with the name in the message.

use vybe_runtime::opcode::Op;
use vybe_runtime::Chunk;

use super::adapter_util::{new_object, struct_set};
use vybe_compiler::primitives::class_slots::{
    self, ClassSlot, ObjSource, PlainNames, ValueSource,
};

/// Build (once per name) a chunk that raises `NotImplementedError`, and leave
/// a reference to it on the stack — the value a bare `module.name` read is.
fn push_raising_callable(chunks: &mut Vec<Chunk>, current: usize, what: &str, line: u32) {
    let idx = {
        let helper = Chunk::new(&format!("__py_unsupported_{}", what.replace('.', "_")));
        chunks.push(helper);
        let at = chunks.len() - 1;
        chunks[at].arity = 0;
        emit_raise(chunks, at, what, line);
        chunks[at].emit_op(Op::RETURN, line);
        at
    };
    chunks[current].emit_op_u16(Op::REF_FUNC, idx as u16, line);
    chunks[current].emit(0, line);
}

/// `raise NotImplementedError("<what> …")` — through the same Python
/// exception construction every `raise` in the language uses, so the object
/// is a real `NotImplementedError` with a working `str()` and `except` match.
fn emit_raise(chunks: &mut [Chunk], current: usize, what: &str, line: u32) {
    let message = format!("{what} is not implemented by this Python runtime");
    chunks[current].emit_string_const(&message, line);
    crate::emitter::runtime_adapter::emit_py_raise(
        chunks,
        current,
        1,
        "NotImplementedError",
        line,
    );
    // `emit_py_raise` leaves a null behind for the value position it replaces;
    // the throw above means it is never observed, but the stack has to balance.
}

/// `python.surface.<module>.<name>` — a function-shaped surface.
pub fn emit_function_surface(
    chunks: &mut Vec<Chunk>,
    current: usize,
    argc: u8,
    what: &str,
    line: u32,
) {
    if argc == 0 {
        push_raising_callable(chunks, current, what, line);
        return;
    }
    for _ in 0..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    emit_raise(chunks, current, what, line);
}

/// `python.typeobj.<module>.<Name>` — a type-shaped surface. The read yields
/// the type object; construction raises.
pub fn emit_type_surface(
    chunks: &mut Vec<Chunk>,
    current: usize,
    argc: u8,
    what: &str,
    line: u32,
) {
    if argc > 0 {
        for _ in 0..argc {
            chunks[current].emit_op(Op::DROP, line);
        }
        emit_raise(chunks, current, what, line);
        return;
    }
    let leaf = what.rsplit('.').next().unwrap_or(what);
    let chunk = &mut chunks[current];
    new_object(chunk, line);
    chunk.emit_dup(line);
    chunk.emit_string_const("type", line);
    let cs_id = class_slots::resolve(&ClassSlot::TypeIdentity, &PlainNames);
    class_slots::emit_class_set(chunk, ObjSource::Stack, &cs_id, ValueSource::Stack, line);
    chunk.emit_dup(line);
    chunk.emit_string_const(leaf, line);
    struct_set(chunk, &ClassSlot::internal("__name__"), line);
    chunk.emit_dup(line);
    chunk.emit_string_const(what, line);
    struct_set(chunk, &ClassSlot::internal("__qualname__"), line);
}
