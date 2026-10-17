//! WinForms ControlBindingsCollection backed by BindingSource and DOM controls.

use vybe_compiler::primitives::class_slots::{self, Dest, ObjSource};
use vybe_compiler::primitives::instructions::core_wasm;
use vybe_compiler::primitives::{collections, globals, loops, ops, strings};
use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;
use vybe_runtime::opcode::heaptype::HT_EXTERN;

use super::bindingsource_adapter;
use super::object_fields::field_slot;

const CONTROL_BINDINGS: &str = "__dotnet_control_bindings";
const SOURCE_BINDINGS: &str = "__bindings";

#[cfg(test)]
mod registry_tests {
    use super::*;
    use std::sync::Arc;
    use vybe_runtime::{Value, VM};

    fn read_registry(vm: &mut VM) -> Value {
        let mut entry = Chunk::new("registry-regression");
        registry(&mut entry, 17);
        entry.emit_op(Op::RETURN, 17);
        let mut chunks = vec![entry];
        globals::declare_free_globals(&mut chunks);
        globals::normalize_global_table(&mut chunks);
        vm.run(chunks).unwrap()
    }

    #[test]
    fn binding_registry_initializes_missing_null_and_undefined_once() {
        for initial in [None, Some(Value::Null), Some(Value::Undefined)] {
            let mut vm = VM::new();
            vybe_compiler::primitives::platforms::register_platforms_all(&mut vm);
            if let Some(value) = initial {
                vm.set_global(CONTROL_BINDINGS, value);
            }
            let Value::Object(first) = read_registry(&mut vm) else {
                panic!("missing registry must initialize an Object");
            };
            let Value::Object(second) = read_registry(&mut vm) else {
                panic!("registry must retain its Object");
            };
            assert!(Arc::ptr_eq(&first, &second));
            assert!(vm.last_exception.is_none());
        }
    }

    #[test]
    fn registry_preserves_nonmissing_values_and_wasm_null_distinction() {
        for initial in [Value::Bool(false), Value::I32(0), Value::String("".into())] {
            let mut vm = VM::new();
            vybe_compiler::primitives::platforms::register_platforms_all(&mut vm);
            vm.set_global(CONTROL_BINDINGS, initial.clone());
            assert_eq!(read_registry(&mut vm), initial);
        }
        let mut vm = VM::new();
        let mut entry = Chunk::new("strict-null-regression");
        vybe_compiler::primitives::expressions::emit_undefined(&mut entry, 17);
        entry.emit_op(Op::REF_IS_NULL, 17);
        entry.emit_op(Op::RETURN, 17);
        vybe_compiler::primitives::platforms::register_platforms_all(&mut vm);
        assert_eq!(vm.run(vec![entry]).unwrap(), Value::I32(0));
    }
}

fn document(chunk: &mut Chunk, line: u32) {
    let active = chunk.add_import("web:html", "activeDocument");
    chunk.emit_call(active, 0, line);
}

fn object_get(chunk: &mut Chunk, object: u16, name: &str, line: u32) -> u16 {
    let slot = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_GET, object, line);
    chunk.emit_string_const(name, line);
    let get = chunk.add_import("ecma:object", "get");
    chunk.emit_call(get, 2, line);
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
    slot
}

fn object_set(chunk: &mut Chunk, object: u16, name: &str, value: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, object, line);
    chunk.emit_string_const(name, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    let set = chunk.add_import("ecma:object", "set");
    chunk.emit_call(set, 3, line);
    chunk.emit_op(Op::DROP, line);
}

fn source_bindings(chunk: &mut Chunk, source: u16, line: u32) -> u16 {
    let slot = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_GET, source, line);
    class_slots::emit_class_get(
        chunk,
        ObjSource::Stack,
        &field_slot(SOURCE_BINDINGS),
        Dest::Stack,
        line,
    );
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
    slot
}

fn control_id(chunk: &mut Chunk, control: u16, line: u32) -> u16 {
    let slot = chunk.alloc_scratch(1);
    document(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, control, line);
    chunk.emit_string_const("id", line);
    let get = chunk.add_import("web:dom", "getAttribute");
    chunk.emit_call(get, 3, line);
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
    slot
}

fn checkable_input(chunk: &mut Chunk, control: u16, line: u32) {
    document(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, control, line);
    let first = chunk.add_import("web:dom", "firstChild");
    chunk.emit_call(first, 2, line);
}

fn registry(chunk: &mut Chunk, line: u32) {
    globals::emit_read(chunk, CONTROL_BINDINGS, line);
    core_wasm::dup(chunk, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    // An uninitialized global reads Undefined, which is a non-null externref.
    // Reuse DUP's existing root rather than reading the registry twice.
    chunk.emit_op_u16(Op::LOCAL_GET, chunk.dup_slot.expect("dup reserved its root"), line);
    let undefined = chunk.add_import("wasm:js-undefined", "test");
    chunk.emit_call(undefined, 1, line);
    chunk.emit_op(Op::I32_OR, line);
    chunk.emit_if_params(line, 1, 1);
    chunk.emit_op(Op::DROP, line);
    let new = chunk.add_import("ecma:object", "new");
    chunk.emit_call(new, 0, line);
    core_wasm::dup(chunk, line);
    globals::emit_write(chunk, CONTROL_BINDINGS, line);
    chunk.emit_end(line);
}

fn same(chunk: &mut Chunk, value: u16, literal: &str, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    chunk.emit_string_const(literal, line);
    ops::emit_dyn_eq(chunk, line);
}

fn row_member_key(chunk: &mut Chunk, row: u16, member: u16, line: u32) -> u16 {
    let key = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_GET, row, line);
    chunk.emit_op_u16(Op::LOCAL_GET, member, line);
    vybe_compiler::primitives::reflection::emit_has_in_in_chunk(chunk, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::LOCAL_GET, member, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, member, line);
    let lower = chunk.add_import("ecma:string", "toLowerCase");
    chunk.emit_call(lower, 1, line);
    chunk.emit_end(line);
    chunk.emit_op_u16(Op::LOCAL_SET, key, line);
    key
}

fn current_row(chunks: &mut [Chunk], current: usize, source: u16, line: u32) -> u16 {
    chunks[current].emit_op_u16(Op::LOCAL_GET, source, line);
    bindingsource_adapter::emit_bindingsource_current(chunks, current, line);
    let row = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, row, line);
    row
}

fn row_value(chunk: &mut Chunk, row: u16, member: u16, line: u32) -> u16 {
    let value = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_GET, row, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if_value(line);
    chunk.emit_ref_null(HT_EXTERN, line);
    chunk.emit_else(line);
    same(chunk, member, "", line);
    chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::LOCAL_GET, row, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, row, line);
    let key = row_member_key(chunk, row, member, line);
    chunk.emit_op_u16(Op::LOCAL_GET, key, line);
    vybe_compiler::primitives::reflection::emit_get_property_in_chunk(chunk, line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_op_u16(Op::LOCAL_SET, value, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if(line);
    chunk.emit_string_const("", line);
    chunk.emit_op_u16(Op::LOCAL_SET, value, line);
    chunk.emit_end(line);
    value
}

fn set_control(chunk: &mut Chunk, control: u16, property: u16, value: u16, line: u32) {
    same(chunk, property, "Checked", line);
    chunk.emit_if(line);
    document(chunk, line);
    checkable_input(chunk, control, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    let boolean = chunk.add_import("ecma:boolean", "toBoolean");
    chunk.emit_call(boolean, 1, line);
    let checked = chunk.add_import("web:html", "setChecked");
    chunk.emit_call(checked, 3, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_else(line);

    same(chunk, property, "Text", line);
    same(chunk, property, "Value", line);
    chunk.emit_op(Op::I32_OR, line);
    chunk.emit_if(line);
    document(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, control, line);
    let tag = chunk.add_import("web:dom", "nodeName");
    chunk.emit_call(tag, 2, line);
    let tag_slot = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, tag_slot, line);
    for (index, name) in ["INPUT", "TEXTAREA", "SELECT"].iter().enumerate() {
        same(chunk, tag_slot, name, line);
        if index != 0 {
            chunk.emit_op(Op::I32_OR, line);
        }
    }
    chunk.emit_if(line);
    document(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, control, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    strings::emit_to_string(chunk, line);
    let set = chunk.add_import("web:html", "setValue");
    chunk.emit_call(set, 3, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_else(line);
    document(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, control, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    strings::emit_to_string(chunk, line);
    let set = chunk.add_import("web:dom", "setTextContent");
    chunk.emit_call(set, 3, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_end(line);
    chunk.emit_else(line);

    document(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, control, line);
    same(chunk, property, "ImageLocation", line);
    chunk.emit_if_value(line);
    chunk.emit_string_const("src", line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, property, line);
    chunk.emit_end(line);
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    strings::emit_to_string(chunk, line);
    let set = chunk.add_import("web:dom", "setAttribute");
    chunk.emit_call(set, 4, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_end(line);
    chunk.emit_end(line);
}

/// Re-read the current row into every control bound to a BindingSource.
pub fn refresh(chunks: &mut [Chunk], current: usize, source: u16, line: u32) {
    refresh_controls(chunks, current, source, line);
    super::datagrid_binding_adapter::refresh_source_grids(chunks, current, source, line);
}

pub(super) fn refresh_controls(chunks: &mut [Chunk], current: usize, source: u16, line: u32) {
    let bindings = source_bindings(&mut chunks[current], source, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, bindings, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    chunks[current].emit_else(line);
    let row = current_row(chunks, current, source, line);
    let len = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_GET, bindings, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len, line);
    let index = chunks[current].alloc_scratch(1);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, index, line);
    let loop_state = loops::emit_loop_start(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, index, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    loops::emit_loop_cond(chunks, current, line);
    let c = &mut chunks[current];
    c.emit_op_u16(Op::LOCAL_GET, bindings, line);
    c.emit_op_u16(Op::LOCAL_GET, index, line);
    c.emit_op(Op::ARRAY_GET, line);
    let binding = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, binding, line);
    let control = object_get(c, binding, "control", line);
    let property = object_get(c, binding, "property", line);
    let member = object_get(c, binding, "member", line);
    let value = row_value(c, row, member, line);
    set_control(c, control, property, value, line);
    c.emit_op_u16(Op::LOCAL_GET, index, line);
    c.emit_i32_const(1, line);
    c.emit_op(Op::I32_ADD, line);
    c.emit_op_u16(Op::LOCAL_SET, index, line);
    loops::emit_loop_end(chunks, current, loop_state, line);
    chunks[current].emit_end(line);
}

fn change_callback(chunks: &mut Vec<Chunk>, line: u32) -> usize {
    let mut body =
        vybe_compiler::primitives::functions::create_function_chunk("__dotnet_binding_change", 1);
    body.local_count = 1;
    let c = &mut body;
    let event = object_get(c, 0, "currentTarget", line);
    let new = c.add_import("ecma:object", "new");
    c.emit_call(new, 0, line);
    let control = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, control, line);
    object_set(c, control, "__node", event, line);
    let id = control_id(c, control, line);
    globals::emit_read(c, CONTROL_BINDINGS, line);
    c.emit_op_u16(Op::LOCAL_GET, id, line);
    let get = c.add_import("ecma:object", "get");
    c.emit_call(get, 2, line);
    let bindings = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, bindings, line);
    c.emit_op_u16(Op::LOCAL_GET, bindings, line);
    c.emit_op(Op::REF_IS_NULL, line);
    c.emit_if(line);
    c.emit_else(line);
    let len = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_GET, bindings, line);
    c.emit_op(Op::ARRAY_LENGTH, line);
    c.emit_op_u16(Op::LOCAL_SET, len, line);
    let index = c.alloc_scratch(1);
    c.emit_i32_const(0, line);
    c.emit_op_u16(Op::LOCAL_SET, index, line);
    let mut body_vec = vec![body];
    let loop_state = loops::emit_loop_start(&mut body_vec, 0, line);
    let c = &mut body_vec[0];
    c.emit_op_u16(Op::LOCAL_GET, index, line);
    c.emit_op_u16(Op::LOCAL_GET, len, line);
    c.emit_op(Op::I32_LT_S, line);
    loops::emit_loop_cond(&mut body_vec, 0, line);
    let c = &mut body_vec[0];
    c.emit_op_u16(Op::LOCAL_GET, bindings, line);
    c.emit_op_u16(Op::LOCAL_GET, index, line);
    c.emit_op(Op::ARRAY_GET, line);
    let binding = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, binding, line);
    let source = object_get(c, binding, "source", line);
    let member = object_get(c, binding, "member", line);
    let property = object_get(c, binding, "property", line);
    let row = current_row(&mut body_vec, 0, source, line);
    let c = &mut body_vec[0];
    c.emit_op_u16(Op::LOCAL_GET, row, line);
    c.emit_op(Op::REF_IS_NULL, line);
    c.emit_if(line);
    c.emit_else(line);
    same(c, property, "Checked", line);
    c.emit_if_value(line);
    document(c, line);
    checkable_input(c, control, line);
    let checked = c.add_import("web:html", "checked");
    c.emit_call(checked, 2, line);
    c.emit_else(line);
    document(c, line);
    c.emit_op_u16(Op::LOCAL_GET, control, line);
    let value = c.add_import("web:html", "value");
    c.emit_call(value, 2, line);
    c.emit_end(line);
    let value = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, value, line);
    same(c, member, "", line);
    c.emit_op(Op::I32_EQZ, line);
    c.emit_if(line);
    let key = row_member_key(c, row, member, line);
    c.emit_op_u16(Op::LOCAL_GET, row, line);
    c.emit_op_u16(Op::LOCAL_GET, key, line);
    c.emit_op_u16(Op::LOCAL_GET, value, line);
    vybe_compiler::primitives::reflection::emit_set_property_in_chunk(c, line);
    c.emit_op(Op::DROP, line);
    c.emit_end(line);
    c.emit_end(line);
    refresh(&mut body_vec, 0, source, line);
    let c = &mut body_vec[0];
    c.emit_op_u16(Op::LOCAL_GET, index, line);
    c.emit_i32_const(1, line);
    c.emit_op(Op::I32_ADD, line);
    c.emit_op_u16(Op::LOCAL_SET, index, line);
    loops::emit_loop_end(&mut body_vec, 0, loop_state, line);
    let c = &mut body_vec[0];
    c.emit_end(line);
    c.emit_ref_null(HT_EXTERN, line);
    c.emit_op(Op::RETURN, line);
    chunks.push(body_vec.remove(0));
    chunks.len() - 1
}

/// Stack: [control, property, source, member] -> [binding].
pub fn add(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let handler = change_callback(chunks, line);
    let c = &mut chunks[current];
    let member = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, member, line);
    let source = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, source, line);
    let property = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, property, line);
    let control = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, control, line);
    let new = c.add_import("ecma:object", "new");
    c.emit_call(new, 0, line);
    let binding = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, binding, line);
    for (key, value) in [
        ("control", control),
        ("source", source),
        ("property", property),
        ("member", member),
    ] {
        object_set(c, binding, key, value, line);
    }
    let source_items = source_bindings(c, source, line);
    c.emit_op_u16(Op::LOCAL_GET, source_items, line);
    c.emit_op_u16(Op::LOCAL_GET, binding, line);
    let push = c.add_import("ecma:array", "push");
    c.emit_call(push, 2, line);
    c.emit_op(Op::DROP, line);

    registry(c, line);
    let global = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, global, line);
    let id = control_id(c, control, line);
    let control_items = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_GET, global, line);
    c.emit_op_u16(Op::LOCAL_GET, id, line);
    let has = c.add_import("ecma:object", "hasOwn");
    c.emit_call(has, 2, line);
    c.emit_if(line);
    c.emit_op_u16(Op::LOCAL_GET, global, line);
    c.emit_op_u16(Op::LOCAL_GET, id, line);
    let get = c.add_import("ecma:object", "get");
    c.emit_call(get, 2, line);
    c.emit_op_u16(Op::LOCAL_SET, control_items, line);
    c.emit_else(line);
    collections::emit_array_new(chunks, current, 0, line);
    let c = &mut chunks[current];
    c.emit_op_u16(Op::LOCAL_SET, control_items, line);
    c.emit_op_u16(Op::LOCAL_GET, global, line);
    c.emit_op_u16(Op::LOCAL_GET, id, line);
    c.emit_op_u16(Op::LOCAL_GET, control_items, line);
    let set = c.add_import("ecma:object", "set");
    c.emit_call(set, 3, line);
    c.emit_op(Op::DROP, line);
    for event in ["input", "change"] {
        vybe_compiler::primitives::gui::emit_add_event_listener(c, control, event, handler, line);
    }
    c.emit_end(line);
    c.emit_op_u16(Op::LOCAL_GET, control_items, line);
    c.emit_op_u16(Op::LOCAL_GET, binding, line);
    let push = c.add_import("ecma:array", "push");
    c.emit_call(push, 2, line);
    c.emit_op(Op::DROP, line);
    refresh(chunks, current, source, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, binding, line);
}
