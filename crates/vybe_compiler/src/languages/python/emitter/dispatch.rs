//! Auto-extracted `python.*` dispatch (language-specific routing lives in the
//! language module; the common dispatcher delegates here).

use vybe_bytecode::Chunk;

pub fn dispatch(name: &str, chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) -> bool {
    match name {
        "python.extend" => {
            crate::emitter::python::collections_adapter::emit_extend(chunks, current, line)
        }
        "python.get" => {
            crate::emitter::python::collections_adapter::emit_get(chunks, current, argc, line)
        }
        "python.pop" => {
            crate::emitter::python::collections_adapter::emit_pop(chunks, current, argc, line)
        }
        "python.index" => {
            crate::emitter::python::collections_adapter::emit_index(chunks, current, argc, line)
        }
        "python.from_end" => {
            crate::emitter::python::collections_adapter::emit_from_end(chunks, current, argc, line)
        }
        "python.contains" => {
            crate::emitter::python::collections_adapter::emit_contains(chunks, current, line)
        }
        "python.bytes_wrap" => {
            crate::emitter::python::bytes_adapter::emit_bytes_wrap(chunks, current, argc, line)
        }
        "python.next" => {
            crate::emitter::python::collections_adapter::emit_pynext(chunks, current, argc, line)
        }
        "python.float_repr" => {
            crate::emitter::python::float_adapter::emit_float_repr(chunks, current, argc, line)
        }
        "python.gen_send" => {
            crate::emitter::python::collections_adapter::emit_gen_send(chunks, current, argc, line)
        }
        "python.gen_throw" => {
            crate::emitter::python::collections_adapter::emit_gen_throw(chunks, current, argc, line)
        }
        "python.add" => {
            crate::emitter::python::collections_adapter::emit_add(chunks, current, line)
        }
        "python.remove" => {
            crate::emitter::python::collections_adapter::emit_remove(chunks, current, line)
        }
        "python.discard" => {
            crate::emitter::python::collections_adapter::emit_discard(chunks, current, line)
        }
        "python.copy" => {
            crate::emitter::python::collections_adapter::emit_copy(chunks, current, line)
        }
        "python.update" => {
            crate::emitter::python::collections_adapter::emit_update(chunks, current, line)
        }
        "python.intersection_update" => {
            crate::emitter::python::collections_adapter::emit_intersection_update(
                chunks, current, line,
            )
        }
        "python.difference_update" => {
            crate::emitter::python::collections_adapter::emit_difference_update(
                chunks, current, line,
            )
        }
        "python.symmetric_difference_update" => {
            crate::emitter::python::collections_adapter::emit_symmetric_difference_update(
                chunks, current, line,
            )
        }
        "python.clear" => {
            crate::emitter::python::collections_adapter::emit_clear(chunks, current, line)
        }
        "python.length" => {
            crate::emitter::python::collections_adapter::emit_length(chunks, current, line)
        }
        "python.str" => {
            crate::emitter::python::runtime_adapter::emit_str(chunks, current, argc, line)
        }
        "python.print" => {
            crate::emitter::python::runtime_adapter::emit_print(chunks, current, argc, line)
        }
        "python.pyadd" => {
            crate::emitter::python::runtime_adapter::emit_pyadd(chunks, current, line)
        }
        "python.pymul" => {
            crate::emitter::python::runtime_adapter::emit_pymul(chunks, current, line)
        }
        "python.range" => {
            crate::emitter::python::runtime_adapter::emit_range(chunks, current, argc, line)
        }
        name if crate::emitter::python::runtime_adapter::emit_helper(
            name, chunks, current, argc, line,
        ) => {}

        // ── COBOL adapters ──
        _ => return false,
    }
    true
}
