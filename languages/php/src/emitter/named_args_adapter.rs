use vybe_compiler::primitives::{class_slots, loops, ops, reflection};
use vybe_runtime::{Chunk, opcode::Op};

fn get(c: &mut Chunk, slot: u16, line: u32) {
    c.emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn set(c: &mut Chunk, slot: u16, line: u32) {
    c.emit_op_u16(Op::LOCAL_SET, slot, line);
}

fn read(c: &mut Chunk, array: u16, index: u16, line: u32) {
    get(c, array, line);
    get(c, index, line);
    c.emit_op(Op::ARRAY_GET, line);
}

fn increment(c: &mut Chunk, slot: u16, line: u32) {
    get(c, slot, line);
    c.emit_f64_const(1.0, line);
    c.emit_op(Op::F64_ADD, line);
    set(c, slot, line);
}

fn method(c: &mut Chunk, object: u16, name: &str, line: u32) {
    get(c, object, line);
    c.emit_string_const(name, line);
    c.emit_bool_const(false, line);
    let lookup = c.add_import("ecma:value", "getMethodForCall");
    c.emit_call(lookup, 3, line);
    get(c, object, line);
    c.emit_array_new_fixed(0, 0, line);
    let apply = c.add_import("ecma:function", "apply");
    c.emit_call(apply, 3, line);
}

fn range_start(chunks: &mut [Chunk], current: usize, index: u16, count: u16, line: u32) -> loops::LoopState {
    chunks[current].emit_f64_const(0.0, line);
    set(&mut chunks[current], index, line);
    let state = loops::emit_loop_start(chunks, current, line);
    get(&mut chunks[current], index, line);
    get(&mut chunks[current], count, line);
    chunks[current].emit_op(Op::F64_LT, line);
    loops::emit_loop_cond(chunks, current, line);
    state
}

// Bind against the live method's metadata, not a caller-file signature cache.
// Values are already evaluated in source order; names only select their slots.
pub fn emit_bind_named_args(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc != 3 {
        return;
    }
    let c = &mut chunks[current];
    let values = c.alloc_scratch(1);
    let names = c.alloc_scratch(1);
    let params = c.alloc_scratch(1);
    let count = c.alloc_scratch(1);
    let supplied = c.alloc_scratch(1);
    let i = c.alloc_scratch(1);
    let j = c.alloc_scratch(1);
    let param = c.alloc_scratch(1);
    let name = c.alloc_scratch(1);
    let target = c.alloc_scratch(1);
    let out = c.alloc_scratch(1);
    let seen = c.alloc_scratch(1);
    let last = c.alloc_scratch(1);
    let named_seen = c.alloc_scratch(1);
    let variadic = c.alloc_scratch(1);
    set(c, values, line);
    set(c, names, line);
    set(c, params, line);
    get(c, params, line);
    let is_array = c.add_import("ecma:array", "isArray");
    c.emit_call(is_array, 1, line);
    ops::emit_dyn_to_bool(c, line);
    c.emit_if_value(line);
    get(c, params, line);
    c.emit_op(Op::ARRAY_LENGTH, line);
    set(c, count, line);
    c.emit_bool_const(false, line);
    set(c, variadic, line);

    let scan = range_start(chunks, current, i, count, line);
    let c = &mut chunks[current];
    read(c, params, i, line);
    set(c, param, line);
    method(c, param, "isvariadic", line);
    ops::emit_dyn_to_bool(c, line);
    c.emit_if(line);
    c.emit_bool_const(true, line);
    set(c, variadic, line);
    c.emit_end(line);
    increment(c, i, line);
    loops::emit_loop_end(chunks, current, scan, line);

    let c = &mut chunks[current];
    get(c, variadic, line);
    ops::emit_dyn_to_bool(c, line);
    c.emit_if_value(line);
    // Keep the pre-existing variadic path until keyed rest binding is available.
    get(c, values, line);
    c.emit_else(line);
    c.emit_array_new_fixed(0, 0, line);
    set(c, out, line);
    c.emit_array_new_fixed(0, 0, line);
    set(c, seen, line);
    c.emit_f64_const(0.0, line);
    set(c, last, line);
    c.emit_bool_const(false, line);
    set(c, named_seen, line);
    get(c, values, line);
    c.emit_op(Op::ARRAY_LENGTH, line);
    set(c, supplied, line);

    let arguments = range_start(chunks, current, i, supplied, line);
    let c = &mut chunks[current];
    read(c, names, i, line);
    set(c, name, line);
    get(c, name, line);
    reflection::emit_typeof_is(c, "string", line);
    c.emit_if(line);
    c.emit_bool_const(true, line);
    set(c, named_seen, line);
    c.emit_f64_const(-1.0, line);
    set(c, target, line);
    let search = range_start(chunks, current, j, count, line);
    let c = &mut chunks[current];
    read(c, params, j, line);
    class_slots::emit_class_get(c, class_slots::ObjSource::Stack,
        &class_slots::resolve(&class_slots::ClassSlot::Internal("name".into()), &class_slots::PlainNames),
        class_slots::Dest::Stack, line);
    get(c, name, line);
    ops::emit_js_strict_eq(c, line);
    ops::emit_dyn_to_bool(c, line);
    c.emit_if(line);
    get(c, j, line);
    set(c, target, line);
    c.emit_end(line);
    increment(c, j, line);
    loops::emit_loop_end(chunks, current, search, line);
    let c = &mut chunks[current];
    get(c, target, line);
    c.emit_f64_const(0.0, line);
    c.emit_op(Op::F64_LT, line);
    c.emit_if(line);
    super::type_guard::emit_throw_const(chunks, current, "Error", "Unknown named parameter", line);
    let c = &mut chunks[current];
    c.emit_end(line);
    c.emit_else(line);
    get(c, named_seen, line);
    ops::emit_dyn_to_bool(c, line);
    c.emit_if(line);
    super::type_guard::emit_throw_const(chunks, current, "Error", "Cannot use positional argument after named argument", line);
    let c = &mut chunks[current];
    c.emit_end(line);
    get(c, i, line);
    set(c, target, line);
    c.emit_end(line);
    read(c, seen, target, line);
    ops::emit_dyn_to_bool(c, line);
    c.emit_if(line);
    super::type_guard::emit_throw_const(chunks, current, "Error", "Named parameter overwrites previous argument", line);
    let c = &mut chunks[current];
    c.emit_end(line);
    get(c, out, line);
    get(c, target, line);
    read(c, values, i, line);
    c.emit_op(Op::ARRAY_SET, line);
    get(c, seen, line);
    get(c, target, line);
    c.emit_bool_const(true, line);
    c.emit_op(Op::ARRAY_SET, line);
    get(c, target, line);
    get(c, last, line);
    c.emit_op(Op::F64_GE, line);
    c.emit_if(line);
    get(c, target, line);
    c.emit_f64_const(1.0, line);
    c.emit_op(Op::F64_ADD, line);
    set(c, last, line);
    c.emit_end(line);
    increment(c, i, line);
    loops::emit_loop_end(chunks, current, arguments, line);

    let defaults = range_start(chunks, current, i, count, line);
    let c = &mut chunks[current];
    read(c, seen, i, line);
    ops::emit_dyn_not(c, line);
    c.emit_if(line);
    read(c, params, i, line);
    set(c, param, line);
    method(c, param, "isoptional", line);
    ops::emit_dyn_not(c, line);
    c.emit_if(line);
    super::type_guard::emit_throw_const(chunks, current, "ArgumentCountError", "Required argument was not supplied", line);
    let c = &mut chunks[current];
    c.emit_end(line);
    get(c, i, line);
    get(c, last, line);
    c.emit_op(Op::F64_LT, line);
    c.emit_if(line);
    get(c, out, line);
    get(c, i, line);
    method(c, param, "getdefaultvalue", line);
    c.emit_op(Op::ARRAY_SET, line);
    c.emit_end(line);
    c.emit_end(line);
    increment(c, i, line);
    loops::emit_loop_end(chunks, current, defaults, line);
    let c = &mut chunks[current];
    get(c, out, line);
    c.emit_end(line);
    c.emit_else(line);
    get(c, values, line);
    c.emit_end(line);
}
