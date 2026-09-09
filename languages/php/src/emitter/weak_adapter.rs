//! PHP WeakReference adapter.
//!
//! Registered as a `php.*` tree leaf and emitted as bytecode over the ECMA
//! weak reference surface. This keeps WeakReference out of the miscellaneous
//! adapter bucket without changing the public PHP resolver name.

use vybe_compiler::primitives::class_slots::{
    self, ClassSlot, Dest, ObjSource, PlainNames, ValueSource,
};
use vybe_runtime::opcode::Op;
use vybe_runtime::Chunk;

/// `WeakReference::create($obj)` -> object with `get()` that derefs the weak ref.
pub fn emit_weak_ref_create(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let get_idx = {
        let mut c = Chunk::new("__weak_ref_get");
        c.arity = 1;
        let cs_slot = class_slots::resolve(&ClassSlot::Internal("__weak".to_string()), &PlainNames);
        c.emit_op_u16(Op::LOCAL_GET, 0, line);
        class_slots::emit_class_get(&mut c, ObjSource::Stack, &cs_slot, Dest::Stack, line);
        let idx = c.add_import("ecma:weakref", "deref");
        c.emit_call(idx, 1, line);
        c.emit_op(Op::RETURN, line);
        c.local_count = c.local_count.max(1);
        chunks.push(c);
        chunks.len() - 1
    };

    let chunk = &mut chunks[current];
    let obj_slot = chunk.alloc_scratch(1);
    let this_slot = chunk.alloc_scratch(1);

    chunk.emit_op_u16(Op::LOCAL_SET, obj_slot, line);

    class_slots::emit_class_alloc(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_SET, this_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    let idx = chunk.add_import("ecma:weakref", "new");
    chunk.emit_call(idx, 1, line);
    let cs_slot = class_slots::resolve(&ClassSlot::Internal("__weak".to_string()), &PlainNames);
    class_slots::emit_class_set(chunk, ObjSource::Stack, &cs_slot, ValueSource::Stack, line);

    chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
    chunk.emit_op_u16(Op::REF_FUNC, get_idx as u16, line);
    chunk.emit(0, line);
    let cs_slot = class_slots::resolve(&ClassSlot::Internal("get".to_string()), &PlainNames);
    class_slots::emit_class_set(chunk, ObjSource::Stack, &cs_slot, ValueSource::Stack, line);

    chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
}
