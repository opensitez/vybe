//! Auto-extracted `python.*` dispatch (language-specific routing lives in the
//! language module; the common dispatcher delegates here).

use vybe_bytecode::Chunk;

pub fn dispatch(name: &str, chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) -> bool {
    match name {
        "python.extend" => {
            crate::emitter::collections_adapter::emit_extend(chunks, current, line)
        }
        "python.get" => {
            crate::emitter::collections_adapter::emit_get(chunks, current, argc, line)
        }
        "python.pop" => {
            crate::emitter::collections_adapter::emit_pop(chunks, current, argc, line)
        }
        "python.index" => {
            crate::emitter::collections_adapter::emit_index(chunks, current, argc, line)
        }
        "python.from_end" => {
            crate::emitter::collections_adapter::emit_from_end(chunks, current, argc, line)
        }
        "python.contains" => {
            crate::emitter::collections_adapter::emit_contains(chunks, current, line)
        }
        "python.next" => {
            crate::emitter::collections_adapter::emit_pynext(chunks, current, argc, line)
        }
        "python.float_repr" => {
            crate::emitter::float_adapter::emit_float_repr(chunks, current, argc, line)
        }
        "python.gen_send" => {
            crate::emitter::collections_adapter::emit_gen_send(chunks, current, argc, line)
        }
        "python.gen_throw" => {
            crate::emitter::collections_adapter::emit_gen_throw(chunks, current, argc, line)
        }
        "python.add" => {
            crate::emitter::collections_adapter::emit_add(chunks, current, line)
        }
        "python.remove" => {
            crate::emitter::collections_adapter::emit_remove(chunks, current, line)
        }
        "python.discard" => {
            crate::emitter::collections_adapter::emit_discard(chunks, current, line)
        }
        "python.copy" => {
            crate::emitter::collections_adapter::emit_copy(chunks, current, line)
        }
        "python.update" => {
            crate::emitter::collections_adapter::emit_update(chunks, current, line)
        }
        "python.intersection_update" => {
            crate::emitter::collections_adapter::emit_intersection_update(
                chunks, current, line,
            )
        }
        "python.difference_update" => {
            crate::emitter::collections_adapter::emit_difference_update(
                chunks, current, line,
            )
        }
        "python.symmetric_difference_update" => {
            crate::emitter::collections_adapter::emit_symmetric_difference_update(
                chunks, current, line,
            )
        }
        "python.clear" => {
            crate::emitter::collections_adapter::emit_clear(chunks, current, line)
        }
        "python.length" => {
            crate::emitter::collections_adapter::emit_length(chunks, current, line)
        }
        "python.str" => {
            crate::emitter::runtime_adapter::emit_str(chunks, current, argc, line)
        }
        "python.print" => {
            crate::emitter::runtime_adapter::emit_print(chunks, current, argc, line)
        }
        "python.bytes_decode" => {
            crate::emitter::runtime_adapter::emit_bytes_decode(chunks, current, argc, line)
        }
        "python.pyadd" => {
            crate::emitter::runtime_adapter::emit_pyadd(chunks, current, line)
        }
        "python.pymul" => {
            crate::emitter::runtime_adapter::emit_pymul(chunks, current, line)
        }
        "python.pysub" => {
            crate::emitter::runtime_adapter::emit_pysub(chunks, current, line)
        }
        "python.pytruediv" => {
            crate::emitter::runtime_adapter::emit_pytruediv(chunks, current, line)
        }
        "python.pyfloordiv" => {
            crate::emitter::runtime_adapter::emit_pyfloordiv(chunks, current, line)
        }
        "python.pymod" => {
            crate::emitter::runtime_adapter::emit_pymod(chunks, current, line)
        }
        "python.pypow" => {
            crate::emitter::runtime_adapter::emit_pypow(chunks, current, line)
        }
        "python.range" => {
            crate::emitter::runtime_adapter::emit_range(chunks, current, argc, line)
        }
        name if crate::emitter::runtime_adapter::emit_helper(
            name, chunks, current, argc, line,
        ) => {}

        // ── COBOL adapters ──
        _ => return false,
    }
    true
}
