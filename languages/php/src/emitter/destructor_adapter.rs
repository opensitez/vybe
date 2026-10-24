//! Guest half of PHP object destruction (host half: `crate::destructors`).
//!
//! Construction registers an object whose Destructor slot is callable; a
//! released value that is not itself such an object is reported only while
//! the `__php_destructibles` flag says a registered object is alive.
use vybe_compiler::primitives::{class_slots, globals, instructions::recipes};
use vybe_runtime::opcode::heaptype::{HT_I31, HeapType};
use vybe_runtime::{Chunk, opcode::Op};

use crate::destructors::LIVE_FLAG;

fn destructor_slot(chunk: &mut Chunk) -> class_slots::ResolvedSlot {
    class_slots::resolve_interned(
        chunk,
        &class_slots::ClassSlot::Slot(vybe_ast::ProtocolSlot::Destructor),
        &class_slots::PlainNames,
    )
}

/// Stack: `[object] -> [object]`. Registers a newly constructed object whose
/// class has a callable destructor.
pub fn emit_track_construction(chunk: &mut Chunk, line: u32) {
    let object = chunk.alloc_scratch(1);
    let destructor = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_TEE, object, line);
    let slot = destructor_slot(chunk);
    class_slots::emit_class_get(chunk, class_slots::ObjSource::Stack, &slot, class_slots::Dest::Stack, line);
    chunk.emit_op_u16(Op::LOCAL_TEE, destructor, line);
    recipes::is_func(chunk, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::LOCAL_GET, object, line);
    chunk.emit_op_u16(Op::LOCAL_GET, destructor, line);
    let track = chunk.add_import("php:object", "track");
    chunk.emit_call(track, 2, line);
    globals::emit_write(chunk, LIVE_FLAG, line);
    chunk.emit_end(line);
    chunk.emit_op_u16(Op::LOCAL_GET, object, line);
}

/// Spill a clone's source so [`emit_track_clone`] can read the destructor
/// the copy shares with it. Stack: `[source, extra…] -> [source, extra…]`,
/// where `extra` is the clone operation's other `argc - 1` operands.
pub fn emit_clone_source_spill(chunk: &mut Chunk, argc: u8, line: u32) -> u16 {
    let extra: Vec<u16> = (1..argc).map(|_| chunk.alloc_scratch(1)).collect();
    for slot in extra.iter().rev() {
        chunk.emit_op_u16(Op::LOCAL_SET, *slot, line);
    }
    let source = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_TEE, source, line);
    for slot in &extra {
        chunk.emit_op_u16(Op::LOCAL_GET, *slot, line);
    }
    source
}

/// Stack: `[copy] -> [copy]`. Registers a clone under its source's
/// destructor: the copy is built without a constructor, so the source is
/// where the bound protocol slot lives.
pub fn emit_track_clone(chunk: &mut Chunk, source: u16, line: u32) {
    let copy = chunk.alloc_scratch(1);
    let destructor = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, copy, line);
    chunk.emit_op_u16(Op::LOCAL_GET, source, line);
    let slot = destructor_slot(chunk);
    class_slots::emit_class_get(chunk, class_slots::ObjSource::Stack, &slot, class_slots::Dest::Stack, line);
    chunk.emit_op_u16(Op::LOCAL_TEE, destructor, line);
    recipes::is_func(chunk, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::LOCAL_GET, copy, line);
    chunk.emit_op_u16(Op::LOCAL_GET, destructor, line);
    let track = chunk.add_import("php:object", "track");
    chunk.emit_call(track, 2, line);
    globals::emit_write(chunk, LIVE_FLAG, line);
    chunk.emit_end(line);
    chunk.emit_op_u16(Op::LOCAL_GET, copy, line);
}

/// Stack: `[value] -> [value]`. Registers every destructible object in a
/// value a host operation built without running a class constructor.
pub fn emit_track_value(chunk: &mut Chunk, line: u32) {
    let value = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_TEE, value, line);
    let track = chunk.add_import("php:object", "trackTree");
    chunk.emit_call(track, 1, line);
    globals::emit_write(chunk, LIVE_FLAG, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
}

/// Stack: `[old, destructor] -> []` — the old value is an object with a
/// callable destructor.
pub fn emit_finalize(chunk: &mut Chunk, line: u32) {
    let finalize = chunk.add_import("php:object", "finalize");
    chunk.emit_call(finalize, 2, line);
    chunk.emit_op(Op::DROP, line);
}

/// Push an i32: is any registered destructible object alive? While it is
/// not, no release can destruct anything.
pub fn emit_live_gate(chunk: &mut Chunk, line: u32) {
    globals::emit_read(chunk, LIVE_FLAG, line);
    chunk.emit_ref_type_op(Op::REF_TEST, HeapType::Abstract(HT_I31), line);
}

/// Stack unchanged. Reports the value in `old` as released; the caller has
/// already passed [`emit_live_gate`].
pub fn emit_release_slot(chunk: &mut Chunk, old: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, old, line);
    let release = chunk.add_import("php:object", "release");
    chunk.emit_call(release, 1, line);
    chunk.emit_op(Op::DROP, line);
}

/// Stack unchanged. The full release of the value in `old`, while any
/// registered object is alive: its own destructor when callable, otherwise
/// whatever destructible objects it owns.
pub fn emit_finalize_or_release_slot(chunk: &mut Chunk, old: u16, line: u32) {
    emit_live_gate(chunk, line);
    chunk.emit_if(line);
    let destructor = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_GET, old, line);
    let slot = destructor_slot(chunk);
    class_slots::emit_class_get(chunk, class_slots::ObjSource::Stack, &slot, class_slots::Dest::Stack, line);
    chunk.emit_op_u16(Op::LOCAL_TEE, destructor, line);
    recipes::is_func(chunk, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::LOCAL_GET, old, line);
    chunk.emit_op_u16(Op::LOCAL_GET, destructor, line);
    emit_finalize(chunk, line);
    chunk.emit_else(line);
    emit_release_slot(chunk, old, line);
    chunk.emit_end(line);
    chunk.emit_end(line);
}

const UNSET_DISPLACED: &str = "__php_unset_displaced";

/// `__php_unset_capture($array, $key)` before `unset($array[$key])`: holds
/// the element being removed in a compiler slot, outside the program roots,
/// while any registered object is alive. Stack: `[array, key] -> [null]`.
pub fn emit_unset_capture(chunk: &mut Chunk, line: u32) {
    let key = chunk.alloc_scratch(1);
    let array = chunk.alloc_scratch(1);
    let displaced = chunk.alloc_scratch(1);
    chunk.local_names.push(vybe_runtime::chunk::LocalName::new(
        displaced,
        UNSET_DISPLACED,
        vybe_runtime::chunk::LocalOrigin::Compiler,
    ));
    chunk.emit_op_u16(Op::LOCAL_SET, key, line);
    chunk.emit_op_u16(Op::LOCAL_SET, array, line);
    emit_live_gate(chunk, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::LOCAL_GET, array, line);
    chunk.emit_op_u16(Op::LOCAL_GET, key, line);
    let read = chunk.add_import("php:array", "read");
    chunk.emit_call(read, 2, line);
    chunk.emit_op_u16(Op::LOCAL_SET, displaced, line);
    chunk.emit_end(line);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

/// `__php_unset_release()` after the removal: releases the element the
/// matching capture holds. Stack: `[] -> [null]`.
pub fn emit_unset_release(chunk: &mut Chunk, line: u32) {
    let displaced = chunk.local_names.iter().rev()
        .find(|local| local.name == UNSET_DISPLACED)
        .map(|local| local.slot);
    if let Some(displaced) = displaced {
        emit_live_gate(chunk, line);
        chunk.emit_if(line);
        emit_release_slot(chunk, displaced, line);
        chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        chunk.emit_op_u16(Op::LOCAL_SET, displaced, line);
        chunk.emit_end(line);
    }
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}
