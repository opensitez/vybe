//! Typed properties on element-backed components that are not visual Controls.

use vybe_compiler::primitives::class_slots::{self, Dest, ObjSource, ValueSource};
use vybe_runtime::Chunk;
use vybe_runtime::opcode::heaptype::HT_EXTERN;

use super::object_fields::field_slot;

pub fn emit_get(chunk: &mut Chunk, key: &str, line: u32) {
    class_slots::emit_class_get(chunk, ObjSource::Stack, &field_slot(key), Dest::Stack, line);
}

pub fn emit_set(chunk: &mut Chunk, key: &str, line: u32) {
    class_slots::emit_class_set(
        chunk,
        ObjSource::Stack,
        &field_slot(key),
        ValueSource::Stack,
        line,
    );
    chunk.emit_ref_null(HT_EXTERN, line);
}
