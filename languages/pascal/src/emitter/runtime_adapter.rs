//! Pascal runtime-surface helpers routed via `common:pascal.*`.

use vybe_compiler::primitives::{collections, expressions, instructions::host, ops, sets};
use vybe_runtime::Chunk;
use vybe_runtime::Op;
use vybe_compiler::primitives::class_slots::{
    self,
};

pub fn emit_helper(
    name: &str,
    chunks: &mut Vec<Chunk>,
    current: usize,
    argc: u8,
    line: u32,
) -> bool {
    // `IntToStr` / `FloatToStr` / `X.ToString()` — a PLAIN conversion. This key
    // must stay the plain host call: it is the hot path under every `__vs`, and
    // the slot-probing variant below allocates a scratch local, which shifts
    // the slot indices a nested Pascal function captures by index.
    if name == "pascal.tostring" {
        // The import table is PER-CHUNK (`Chunk::add_import` pushes onto
        // `self.imports`), so the index must come from the chunk the call is
        // emitted into. Resolving it against chunk 0 was invisible at top level
        // — there `current` IS 0 — and inside any function or method the index
        // landed on whatever import happened to occupy that slot, so
        // `FloatToStr(x)` returned `0`, `true` or an output-stream handle
        // depending on the chunk.
        let to_str = chunks[current].add_import("ecma:string", "String");
        chunks[current].emit_call(to_str, argc, line);
        return true;
    }

    // `[builtin_slots.string] to_string` — the binding the concat operator and
    // string interpolation resolve per operand. Separate key from
    // `pascal.tostring` precisely so the probe reaches only the operands of a
    // stringification, never every integer conversion in the program.
    if name == "pascal.to_string_slot" {
        emit_to_string(chunks, current, line);
        return true;
    }

    if name == "pascal.ord" {
        emit_ord(chunks, current, line);
        return true;
    }

    if name == "pascal.succ" || name == "pascal.pred" {
        emit_succ_pred(chunks, current, line, name == "pascal.succ");
        return true;
    }

    if name == "pascal.high" || name == "pascal.low" {
        emit_high_low(chunks, current, line, name == "pascal.high");
        return true;
    }

    if name == "pascal.random" {
        emit_pascal_random(chunks, current, argc, line);
        return true;
    }

    if name == "pascal.assert_message" {
        emit_pascal_assert_message(chunks, current, line);
        return true;
    }

    if name == "pascal.paramstr" {
        chunks[current].emit_op(Op::DROP, line);
        chunks[current].emit_string_const("", line);
        return true;
    }

    if name == "pascal.set_length" {
        sets::emit_size(chunks, current, line);
        return true;
    }

    if name == "pascal.int_xor" {
        chunks[current].emit_op(Op::I32_XOR, line);
        return true;
    }

    if name == "pascal.file_eof" {
        emit_pascal_file_eof(chunks, current, line);
        return true;
    }

    if name == "pascal.string_to_guid" || name == "pascal.guid_to_string" {
        emit_pascal_guid_string(chunks, current, line);
        return true;
    }

    if name == "pascal.is_equal_guid" {
        emit_pascal_guid_eq(chunks, current, line);
        return true;
    }

    if name == "pascal.f64_eq" || name == "pascal.f64_ne" {
        if name == "pascal.f64_eq" {
            chunks[current].emit_op(Op::F64_EQ, line);
        } else {
            chunks[current].emit_op(Op::F64_NE, line);
        }
        vybe_compiler::primitives::ops::emit_i32_to_bool(&mut chunks[current], line);
        return true;
    }

    if name == "pascal.currency_round" {
        emit_currency_round(chunks, current, line);
        return true;
    }

    if name == "pascal.json_stringify" {
        let stringify = chunks[current].add_import("ecma:json", "stringify");
        chunks[current].emit_call(stringify, argc, line);
        return true;
    }

    if name == "pascal.json_parse" {
        vybe_compiler::primitives::json::emit_parse_or_null(chunks, current, line);
        return true;
    }

    if name == "pascal.json_object_new" {
        emit_object_new(chunks, current, line);
        return true;
    }

    if name == "pascal.json_array_new" {
        chunks[current].emit_array_new_fixed(0, 0, line);
        return true;
    }

    if name == "pascal.json_add_pair" {
        collections::emit_set(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        return true;
    }

    if name == "pascal.json_array_add" {
        collections::emit_push(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        return true;
    }

    if name == "pascal.json_count" {
        collections::emit_len(chunks, current, line);
        return true;
    }

    if name == "pascal.json_remove_pair" {
        emit_json_remove_pair(chunks, current, line);
        return true;
    }

    if name == "pascal.json_clone" {
        emit_json_clone(chunks, current, line);
        return true;
    }

    if name == "pascal.json_entries" {
        let entries = chunks[current].add_import("ecma:object", "entries");
        chunks[current].emit_call(entries, argc, line);
        return true;
    }

    if name == "pascal.xml_document_new" {
        emit_xml_document_new(chunks, current, line);
        return true;
    }

    if name == "pascal.xml_load_from_xml" {
        emit_xml_load_from_xml(chunks, current, line);
        return true;
    }

    if name == "pascal.xml_save" {
        emit_xml_save(chunks, current, line);
        return true;
    }

    if name == "pascal.xml_child_node" {
        emit_xml_child_node(chunks, current, line);
        return true;
    }

    if name == "pascal.xml_add_child" {
        emit_xml_add_child(chunks, current, line);
        return true;
    }

    if name == "pascal.xml_set_text" {
        emit_xml_set_text(chunks, current, line);
        return true;
    }

    if name == "pascal.xml_clone_node" {
        let clone = chunks[current].add_import("web:dom-parser", "cloneNode");
        chunks[current].emit_call(clone, argc, line);
        return true;
    }

    if name == "pascal.xml_remove_child" {
        let remove = chunks[current].add_import("web:dom-parser", "removeChild");
        chunks[current].emit_call(remove, argc, line);
        return true;
    }

    let global = match name {
        "pascal.str_remove_range" => "__vybe_pascal_str_remove_range",
        "pascal.str_insert" => "__vybe_pascal_str_insert",
        "pascal.sort_in_place" => "__vybe_sort_in_place",
        _ => return false,
    };
    collections::emit_runtime_helper_call(chunks, current, global, argc, line);
    true
}

fn emit_json_remove_pair(chunks: &mut [Chunk], current: usize, line: u32) {
    let key_slot = chunks[current].alloc_scratch(2);
    let obj_slot = key_slot + 1;
    chunks[current].emit_op_u16(Op::LOCAL_SET, key_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, obj_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key_slot, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    let old_slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, old_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key_slot, line);
    let del = chunks[current].add_import("ecma:object", "delete");
    chunks[current].emit_call(del, 2, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, old_slot, line);
}

fn emit_json_clone(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let stringify = chunks[current].add_import("ecma:json", "stringify");
    chunks[current].emit_call(stringify, 1, line);
    vybe_compiler::primitives::json::emit_parse_or_null(chunks, current, line);
}

/// `TXMLDocument.Create` — stack `[]` → `[document]`.
///
/// The document IS the DOM document. It used to be an `ecma:object` holding
/// one in a `__dom` property, and that box is exactly what stopped a Pascal
/// document from being a PHP `DOMDocument`: the node every other language
/// hands around was one dereference away, and only for the DOCUMENT — Pascal
/// NODES were already raw DOM nodes. `AddChild` on a document read `__dom`
/// while `AddChild` on an element did not, which is where the asymmetry
/// showed up as a failure.
fn emit_xml_document_new(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunks[current].emit_string_const("", line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    let create = chunks[current].add_import("web:dom-parser", "createDocument");
    chunks[current].emit_call(create, 3, line);
}

fn emit_object_new(chunks: &mut [Chunk], current: usize, line: u32) {
    let idx = chunks[current].add_import("ecma:object", "new");
    chunks[current].emit_call(idx, 0, line);
}

/// `doc.LoadFromXML(text)` — stack `[doc, text]` → `[null]`.
///
/// Delphi's is a PROCEDURE: it fills the document the caller already holds,
/// so `doc` has to keep its identity. With the `__dom` box gone there is
/// nothing to reassign, so the parsed root is adopted INTO the document —
/// which is what the DOM says loading is anyway.
fn emit_xml_load_from_xml(chunks: &mut [Chunk], current: usize, line: u32) {
    let xml_slot = chunks[current].alloc_scratch(3);
    let doc_slot = xml_slot + 1;
    let root_slot = xml_slot + 2;
    chunks[current].emit_op_u16(Op::LOCAL_SET, xml_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, doc_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, xml_slot, line);
    vybe_compiler::primitives::xml::emit_parse(chunks, current, 1, line);
    chunks[current].emit_string_const("documentElement", line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, root_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, doc_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, root_slot, line);
    let append = chunks[current].add_import("web:dom-parser", "appendChild");
    chunks[current].emit_call(append, 2, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, doc_slot, line);
    chunks[current].emit_string_const("documentElement", line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, root_slot, line);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

fn emit_xml_save(chunks: &mut [Chunk], current: usize, line: u32) {
    let doc_slot = chunks[current].alloc_scratch(3);
    let xml_slot = doc_slot + 1;
    let version_slot = doc_slot + 2;
    chunks[current].emit_op_u16(Op::LOCAL_SET, doc_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, doc_slot, line);
    vybe_compiler::primitives::xml::emit_save(chunks, current, 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, xml_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, doc_slot, line);
    chunks[current].emit_string_const("Version", line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, version_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, version_slot, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, doc_slot, line);
    chunks[current].emit_string_const("version", line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, version_slot, line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, version_slot, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, xml_slot, line);
    chunks[current].emit_else(line);
    chunks[current].emit_string_const("<?xml version=\"", line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, version_slot, line);
    chunks[current].emit_string_const("\"?>", line);
    let concat = chunks[current].add_import("wasm:js-string", "concat");
    chunks[current].emit_call(concat, 2, line);
    let concat = chunks[current].add_import("wasm:js-string", "concat");
    chunks[current].emit_call(concat, 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, xml_slot, line);
    let concat = chunks[current].add_import("wasm:js-string", "concat");
    chunks[current].emit_call(concat, 2, line);
    chunks[current].emit_end(line);
}

/// `node.ChildNodes['tag']` — stack `[parent, tag]` → `[element_or_null]`.
///
/// Only the BY-NAME case reaches here; the walker leaves an ordinal as a
/// plain index into the shared `childNodes` array. `xml.elements` is the
/// shared `getElementsByTagName`, so the node this hands back is the same
/// node PHP's `DOMDocument` or .NET's `XElement` would.
///
/// Document order means a direct child always precedes its own descendants,
/// so `[0]` is the direct child whenever there is one. It differs from
/// Delphi only when a parent has NO direct child of that name but a deeper
/// one exists — Delphi answers nil, this answers the descendant.
fn emit_xml_child_node(chunks: &mut [Chunk], current: usize, line: u32) {
    vybe_compiler::primitives::xml::emit_elements(chunks, current, 2, line);
    chunks[current].emit_i32_const(0, line);
    collections::emit_get(chunks, current, line);
}

/// `node.AddChild('tag')` — stack `[parent, name]` → `[new_element]`.
///
/// `appendChild` maintains `childNodes`/`parentNode` but NOT
/// `documentElement` — per the DOM that property is set when the root is
/// installed, and `createDocument` is the only host path that does it. So a
/// root appended to a document has to be recorded here, or
/// `doc.DocumentElement` reads the `null` the document was born with.
fn emit_xml_add_child(chunks: &mut [Chunk], current: usize, line: u32) {
    let name_slot = chunks[current].alloc_scratch(3);
    let parent_slot = name_slot + 1;
    let child_slot = name_slot + 2;
    chunks[current].emit_op_u16(Op::LOCAL_SET, name_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, parent_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, name_slot, line);
    let create = chunks[current].add_import("web:dom-parser", "createElement");
    chunks[current].emit_call(create, 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, child_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, parent_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, child_slot, line);
    let append = chunks[current].add_import("web:dom-parser", "appendChild");
    chunks[current].emit_call(append, 2, line);
    chunks[current].emit_op(Op::DROP, line);

    // A DOCUMENT parent gets its root recorded. `nodeType == 9` is the DOM's
    // own discriminator (Living Standard §4.4) — no Pascal-side marker.
    chunks[current].emit_op_u16(Op::LOCAL_GET, parent_slot, line);
    chunks[current].emit_string_const("nodeType", line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_i32_const(9, line);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, parent_slot, line);
    chunks[current].emit_string_const("documentElement", line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, child_slot, line);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, child_slot, line);
}

fn emit_xml_set_text(chunks: &mut [Chunk], current: usize, line: u32) {
    let value_slot = chunks[current].alloc_scratch(2);
    let node_slot = value_slot + 1;
    chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, node_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, node_slot, line);
    chunks[current].emit_string_const("textContent", line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    collections::emit_set(chunks, current, line);
}

fn emit_pascal_guid_string(chunks: &mut [Chunk], current: usize, line: u32) {
    let string_idx = chunks[current].add_import("ecma:string", "String");
    let upper_idx = chunks[current].add_import("ecma:string", "toUpperCase");
    chunks[current].emit_call(string_idx, 1, line);
    chunks[current].emit_call(upper_idx, 1, line);
}

fn emit_pascal_guid_eq(chunks: &mut [Chunk], current: usize, line: u32) {
    let right_slot = chunks[current].alloc_scratch(1);
    let left_slot = chunks[current].alloc_scratch(1);

    chunks[current].emit_op_u16(Op::LOCAL_SET, right_slot, line);
    emit_pascal_guid_string(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, left_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, right_slot, line);
    emit_pascal_guid_string(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, left_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
}

fn emit_pascal_file_eof(chunks: &mut [Chunk], current: usize, line: u32) {
    let handle_slot = chunks[current].alloc_scratch(1);
    let next_slot = chunks[current].alloc_scratch(1);
    let rows_slot = chunks[current].alloc_scratch(1);
    let len_slot = chunks[current].alloc_scratch(1);
    let eof_map_slot = ensure_global_map(chunks, current, "__vb_file_eof_by_handle", line);
    let next_map_slot =
        ensure_global_map(chunks, current, "__vb_record_next_index_by_handle", line);
    let rows_map_slot = ensure_global_map(chunks, current, "__vb_record_rows_by_handle", line);

    {
        let chunk = &mut chunks[current];
        chunk.emit_op_u16(Op::LOCAL_SET, handle_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, next_map_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, handle_slot, line);
        chunk.emit_op(Op::ARRAY_GET, line);
        chunk.emit_op_u16(Op::LOCAL_SET, next_slot, line);

        chunk.emit_op_u16(Op::LOCAL_GET, rows_map_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, handle_slot, line);
        chunk.emit_op(Op::ARRAY_GET, line);
        chunk.emit_op_u16(Op::LOCAL_SET, rows_slot, line);

        chunk.emit_op_u16(Op::LOCAL_GET, next_slot, line);
        chunk.emit_op(Op::REF_IS_NULL, line);
    }

    chunks[current].emit_if(line);
    {
        let chunk = &mut chunks[current];
        chunk.emit_op_u16(Op::LOCAL_GET, eof_map_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, handle_slot, line);
        chunk.emit_op(Op::ARRAY_GET, line);
        vybe_compiler::primitives::instructions::core_wasm::dup(chunk, line);
        chunk.emit_op(Op::REF_IS_NULL, line);
    }
    chunks[current].emit_if(line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_bool_const(false, line);
    chunks[current].emit_else(line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);
    {
        let chunk = &mut chunks[current];
        chunk.emit_op_u16(Op::LOCAL_GET, rows_slot, line);
        chunk.emit_op(Op::REF_IS_NULL, line);
    }
    chunks[current].emit_if(line);
    {
        let chunk = &mut chunks[current];
        chunk.emit_op_u16(Op::LOCAL_GET, eof_map_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, handle_slot, line);
        chunk.emit_op(Op::ARRAY_GET, line);
        vybe_compiler::primitives::instructions::core_wasm::dup(chunk, line);
        chunk.emit_op(Op::REF_IS_NULL, line);
    }
    chunks[current].emit_if(line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_bool_const(false, line);
    chunks[current].emit_else(line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);
    {
        let chunk = &mut chunks[current];
        chunk.emit_op_u16(Op::LOCAL_GET, rows_slot, line);
    }
    collections::emit_len(chunks, current, line);
    {
        let chunk = &mut chunks[current];
        chunk.emit_op_u16(Op::LOCAL_SET, len_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, next_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, len_slot, line);
        vybe_compiler::primitives::ops::emit_dyn_ge(chunk, line);
    }
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

fn ensure_global_map(chunks: &mut [Chunk], current: usize, name: &str, line: u32) -> u16 {
    let slot = chunks[current].alloc_scratch(1);
    vybe_compiler::primitives::globals::emit_read(&mut chunks[current], name, line);
    vybe_compiler::primitives::instructions::core_wasm::dup(&mut chunks[current], line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op(Op::DROP, line);
    collections::emit_map_new(chunks, current, line);
    vybe_compiler::primitives::instructions::core_wasm::dup(&mut chunks[current], line);
    vybe_compiler::primitives::globals::emit_write(&mut chunks[current], name, line);
    chunks[current].emit_else(line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, slot, line);
    slot
}

/// Pascal's string conversion — the `[builtin_slots.string] to_string` binding.
///
/// There are two bindings and this reads BOTH. The profile binding names WHICH
/// conversion Pascal uses; for a class instance the conversion is not a fixed
/// host call but the object's own — so this probes the `ToString` PROTOCOL SLOT
/// (`expressions::emit_rich_to_string`), which is what a Pascal class fills by
/// declaring `function ToString: string` and what a class from ANY language
/// fills with its own spelling (`to_s`, `__str__`, `toString`).
///
/// Reaching it by ROLE is the point: nothing here matches the name `ToString`,
/// so a Pascal `WriteLn(obj)` renders a Ruby or Dart object correctly too.
/// Per flexclassplan §2f-bis an unbound slot falls back to the platform
/// rendering, which is the `else` arm.
///
/// The `typeof`/`isArray` guard is required, not defensive:
/// `emit_rich_to_string` opens with `STRUCT_GET`, which traps on a primitive.
fn emit_to_string(chunks: &mut [Chunk], current: usize, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    host::emit(&mut chunks[current], "ecma:value", "typeof", 1, line);
    chunks[current].emit_string_const("object", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);

    // An array is an object too. Pascal has no array `ToString`, so leave it on
    // the platform coercion rather than probing a slot it can never bind.
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

fn emit_ord(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let value_slot = chunk.local_count;
    let result_slot = value_slot + 1;
    chunk.alloc_scratch(2);

    chunk.emit_op_u16(Op::LOCAL_SET, value_slot, line);

    let str_test = chunk.add_import("wasm:js-string", "test");
    let char_code_at = chunk.add_import("ecma:string", "charCodeAt");
    let bool_test = chunk.add_import("wasm:js-boolean", "test");
    let bool_cast = chunk.add_import("wasm:js-boolean", "cast");
    let num_to_f64 = chunk.add_import("wasm:js-number", "toF64");

    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunk.emit_call(str_test, 1, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunk.emit_i32_const(0, line);
    chunk.emit_call(char_code_at, 2, line);
    chunk.emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunk.emit_call(bool_test, 1, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunk.emit_call(bool_cast, 1, line);
    chunk.emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunk.emit_call(num_to_f64, 1, line);
    chunk.emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunk.emit_end(line);
    chunk.emit_end(line);

    chunk.emit_op_u16(Op::LOCAL_GET, result_slot, line);
}

fn emit_succ_pred(chunks: &mut [Chunk], current: usize, line: u32, increment: bool) {
    let chunk = &mut chunks[current];
    let value_slot = chunk.local_count;
    chunk.alloc_scratch(1);

    let str_test = chunk.add_import("wasm:js-string", "test");
    let char_code_at = chunk.add_import("ecma:string", "charCodeAt");
    let from_char_code = chunk.add_import("ecma:string", "fromCharCode");

    chunk.emit_op_u16(Op::LOCAL_SET, value_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunk.emit_call(str_test, 1, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunk.emit_i32_const(0, line);
    chunk.emit_call(char_code_at, 2, line);
    chunk.emit_i32_const(1, line);
    chunk.emit_op(if increment { Op::I32_ADD } else { Op::I32_SUB }, line);
    chunk.emit_call(from_char_code, 1, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunk.emit_f64_const(1.0, line);
    if increment {
        vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
    } else {
        chunk.emit_op(Op::F64_SUB, line);
    }
    chunk.emit_end(line);
}

fn emit_high_low(chunks: &mut [Chunk], current: usize, line: u32, high: bool) {
    let chunk = &mut chunks[current];
    let value_slot = chunk.local_count;
    chunk.alloc_scratch(1);

    let str_test = chunk.add_import("wasm:js-string", "test");
    let str_length = chunk.add_import("wasm:js-string", "length");

    chunk.emit_op_u16(Op::LOCAL_SET, value_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunk.emit_call(str_test, 1, line);
    chunk.emit_if(line);
    if high {
        chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
        chunk.emit_call(str_length, 1, line);
    } else {
        chunk.emit_i32_const(1, line);
    }
    chunk.emit_else(line);
    if high {
        chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
        chunk.emit_op(Op::ARRAY_LENGTH, line);
        chunk.emit_i32_const(1, line);
        chunk.emit_op(Op::I32_SUB, line);
    } else {
        chunk.emit_i32_const(0, line);
    }
    chunk.emit_end(line);
}

fn emit_pascal_random(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let upper_slot = chunk.local_count;
    chunk.alloc_scratch(1);
    let random = chunk.add_import("ecma:math", "random");
    let floor = chunk.add_import("ecma:math", "floor");

    if argc == 0 {
        chunk.emit_call(random, 0, line);
        return;
    }

    chunk.emit_op_u16(Op::LOCAL_SET, upper_slot, line);
    chunk.emit_call(random, 0, line);
    chunk.emit_op_u16(Op::LOCAL_GET, upper_slot, line);
    chunk.emit_op(Op::F64_MUL, line);
    chunk.emit_call(floor, 1, line);
}

fn emit_pascal_assert_message(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let replace_all = chunk.add_import("ecma:string", "replaceAll");

    chunk.emit_string_const("\r\n", line);
    chunk.emit_string_const("\n", line);
    chunk.emit_call(replace_all, 3, line);
    chunk.emit_string_const("\r", line);
    chunk.emit_string_const("\n", line);
    chunk.emit_call(replace_all, 3, line);
}

fn emit_currency_round(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let value_slot = chunk.local_count;
    chunk.alloc_scratch(1);
    let floor = chunk.add_import("ecma:math", "floor");
    let ceil = chunk.add_import("ecma:math", "ceil");

    chunk.emit_op_u16(Op::LOCAL_SET, value_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunk.emit_f64_const(0.0, line);
    chunk.emit_op(Op::F64_GE, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunk.emit_f64_const(0.5, line);
    chunk.emit_op(Op::F64_ADD, line);
    chunk.emit_call(floor, 1, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunk.emit_f64_const(0.5, line);
    chunk.emit_op(Op::F64_SUB, line);
    chunk.emit_call(ceil, 1, line);
    chunk.emit_end(line);
}

/// Pascal `Write(a, b, c)` / `WriteLn(a, b, c)`.
///
/// Pascal CONCATENATES its arguments with NO separator — `WriteLn('n=', 42)` is
/// `n=42`. Binding these to the shared `print` sent them to
/// `wasi:logging/logging.log`, whose host binding joins with a space and
/// overloads arity as (level, context, message), so vybe printed `n= 42` and
/// dropped middle arguments. Measured against real `fpc`:
/// `WriteLn('x=',i,' y=',i)` is `x=42 y=42`, not `x= 42  y= 42`.
///
/// This is the SAME path php `echo` uses — each value written straight to
/// `wasi:cli/stdout` via `io::emit_write_or_buffer`, one at a time, so output
/// buffering and redirection keep working. No host change, no `wasi:logging`.
pub fn emit_pascal_write(
    chunks: &mut Vec<Chunk>,
    current: usize,
    argc: u8,
    newline: bool,
    line: u32,
) {
    use vybe_runtime::opcode::Op;

    // NOTE: no `strings::emit_concat` here. That helper allocates by bumping
    // `chunk.local_count` directly, while these slots come from
    // `alloc_scratch` — two counters over the same locals. Mixing them
    // clobbered the ENCLOSING function's locals, so a form constructor lost
    // `Self` and reported "undefined is not callable". Each part is written
    // separately instead, which needs no extra locals at all.
    let mut slots = Vec::with_capacity(argc as usize);
    for _ in 0..argc {
        let s = chunks[current].alloc_scratch(1);
        chunks[current].emit_op_u16(Op::LOCAL_SET, s, line);
        slots.push(s);
    }
    slots.reverse();

    // Pascal concatenates its arguments with NO separator. Each part goes
    // through `io::emit_write_or_buffer` — the funnel the doc comment above
    // has always claimed. The code did NOT: it opened its own stream with the
    // WASI 0.2 pair (`get-stdout` + `[method]output-stream.
    // blocking-write-and-flush`), which is a `wasi:io` package 0.3 DELETED,
    // and which reaches the sink WITHOUT consulting the output buffer. That is
    // the defect `io.rs` describes in its own header — a buffer some writers
    // respect and others bypass is not a buffer.
    for s in &slots {
        chunks[current].emit_op_u16(Op::LOCAL_GET, *s, line);
        vybe_compiler::primitives::strings::emit_to_string(&mut chunks[current], line);
        vybe_compiler::primitives::io::emit_write_or_buffer(chunks, current, line);
    }
    if newline {
        chunks[current].emit_string_const("\n", line);
        vybe_compiler::primitives::io::emit_write_or_buffer(chunks, current, line);
    }
}

// ── SysUtils / System stdlib ────────────────────────────────────────────────
//
// These were implemented inline in shared `builtins.rs` behind a
// `profile.name == "pascal"` check. They are pascal's stdlib, so they belong
// here and bind through pascal's profile like any other builtin. Behaviour is
// transcribed unchanged; only the location moved.

fn host(chunks: &mut [Chunk], current: usize, module: &str, name: &str, argc: u8, line: u32) {
    let idx = chunks[current].add_import(module.to_string(), name.to_string());
    chunks[current].emit_call(idx, argc, line);
}

/// `Integer(x)` / `Int(x)` / `LongInt(x)` — truncate toward zero.
pub fn emit_trunc_cast(chunks: &mut [Chunk], current: usize, line: u32) {
    vybe_compiler::primitives::math::emit_trunc(&mut chunks[current], line);
}

/// `IntToHex(v[, width])` — uppercase hex, optionally zero-padded.
pub fn emit_int_to_hex(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    use vybe_runtime::opcode::Op;
    let width = if argc >= 2 {
        let w = chunks[current].alloc_scratch(1);
        chunks[current].emit_op_u16(Op::LOCAL_SET, w, line);
        Some(w)
    } else {
        None
    };
    host(chunks, current, "ecma:number", "Number", 1, line);
    chunks[current].emit_f64_const(16.0, line);
    host(chunks, current, "ecma:number", "toString", 2, line);
    host(chunks, current, "ecma:string", "toUpperCase", 1, line);
    if let Some(w) = width {
        chunks[current].emit_op_u16(Op::LOCAL_GET, w, line);
        chunks[current].emit_string_const("0", line);
        host(chunks, current, "ecma:string", "padStart", 3, line);
    }
}

/// `RGB(r, g, b)` — a colour, as `#RRGGBB`.
///
/// Delphi's `TColor` is a packed integer, but the thing it is ASSIGNED to is a
/// CSS colour, and `#RRGGBB` is the spelling both ends already agree on:
/// `widgets`' `parse_color` reads it, and so does a browser. Packing it
/// into an integer here would mean unpacking it again at every property write,
/// with VCL's byte order (`$00BBGGRR`, blue-first) as a second thing to get
/// wrong.
///
/// Two hex digits per component, zero-padded, reusing `IntToHex` — a colour is
/// not a second hex formatter.
pub fn emit_rgb(chunks: &mut [Chunk], current: usize, line: u32) {
    let blue = chunks[current].alloc_scratch(1);
    let green = chunks[current].alloc_scratch(1);
    let red = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, blue, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, green, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, red, line);

    chunks[current].emit_string_const("#", line);
    for slot in [red, green, blue] {
        chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
        chunks[current].emit_f64_const(2.0, line);
        emit_int_to_hex(chunks, current, 2, line);
        vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
    }
}

/// `BoolToStr(b)` → `true`/`false`; `BoolToStr(b, True)` → `True`/`False`.
pub fn emit_bool_to_str(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc >= 2 {
        chunks[current].emit_op(vybe_runtime::opcode::Op::DROP, line);
    }
    let (t, f) = if argc >= 2 {
        ("True", "False")
    } else {
        ("true", "false")
    };
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_string_const(t, line);
    chunks[current].emit_else(line);
    chunks[current].emit_string_const(f, line);
    chunks[current].emit_end(line);
}

/// `AnsiUpperCase` / `AnsiLowerCase`.
pub fn emit_ansi_case(chunks: &mut [Chunk], current: usize, upper: bool, line: u32) {
    let m = if upper { "toUpperCase" } else { "toLowerCase" };
    host(chunks, current, "ecma:string", m, 1, line);
}

/// `ExtractFileExt(path)` — the extension, **including the dot**.
///
/// The path splitting itself is `wasi:filesystem.pathGetExtension`, which every
/// language's path functions already go through (PHP's `pathinfo` family too) —
/// there is no second path parser here. It answers `pas`, because that is what
/// a path component IS; Delphi answers `.pas`. So the dot is the whole of the
/// adaptation, and it belongs in the language rather than in the host.
///
/// The empty case must stay empty: `ExtractFileExt('README')` is `''`, not
/// `'.'`, which is why this is a branch and not a concatenation.
pub fn emit_extract_file_ext(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];

    let ext = chunk.alloc_scratch(1);

    vybe_compiler::primitives::paths::emit_extension(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_SET, ext, line);

    chunk.emit_op_u16(Op::LOCAL_GET, ext, line);
    chunk.emit_string_const("", line);
    vybe_compiler::primitives::ops::emit_dyn_ne(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    chunk.emit_string_const(".", line);
    chunk.emit_op_u16(Op::LOCAL_GET, ext, line);
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_SET, ext, line);
    chunk.emit_end(line);

    chunk.emit_op_u16(Op::LOCAL_GET, ext, line);
}

/// `SameStr(a, b)` — case-SENSITIVE equality.
pub fn emit_same_str(chunks: &mut [Chunk], current: usize, line: u32) {
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_i32_to_bool(&mut chunks[current], line);
}

/// `SameText(a, b)` / `CompareText(a, b)` — case-INSENSITIVE.
pub fn emit_same_text(chunks: &mut [Chunk], current: usize, equality: bool, line: u32) {
    use vybe_runtime::opcode::Op;
    let b = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, b, line);
    host(chunks, current, "ecma:string", "toLowerCase", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, b, line);
    host(chunks, current, "ecma:string", "toLowerCase", 1, line);
    if equality {
        vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
        vybe_compiler::primitives::ops::emit_i32_to_bool(&mut chunks[current], line);
    } else {
        host(chunks, current, "ecma:string", "localeCompare", 2, line);
    }
}

/// `StrToBool(s)` — only the literal text `true` is true.
pub fn emit_str_to_bool(chunks: &mut [Chunk], current: usize, line: u32) {
    host(chunks, current, "ecma:string", "toLowerCase", 1, line);
    chunks[current].emit_string_const("true", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_i32_to_bool(&mut chunks[current], line);
}

/// `StrToIntDef(s, default)` — parse, falling back when not a number.
pub fn emit_str_to_int_def(chunks: &mut [Chunk], current: usize, line: u32) {
    use vybe_runtime::opcode::Op;
    let dflt = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, dflt, line);
    host(chunks, current, "ecma:number", "parseInt", 1, line);
    let parsed = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, parsed, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, parsed, line);
    host(chunks, current, "ecma:number", "isNaN", 1, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, dflt, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, parsed, line);
    chunks[current].emit_end(line);
}

/// `StrToFloatDef(s, default)` — the float twin of [`emit_str_to_int_def`],
/// differing only in which ECMA parse it delegates to.
pub fn emit_str_to_float_def(chunks: &mut [Chunk], current: usize, line: u32) {
    use vybe_runtime::opcode::Op;
    let dflt = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, dflt, line);
    host(chunks, current, "ecma:number", "parseFloat", 1, line);
    let parsed = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, parsed, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, parsed, line);
    host(chunks, current, "ecma:number", "isNaN", 1, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, dflt, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, parsed, line);
    chunks[current].emit_end(line);
}

/// Pascal `Delete(var S; Index; Count)` — stack `[s, index, count]`.
///
/// Pascal string positions are 1-BASED, so the index is decremented before the
/// shared splice. Nothing is inserted.
pub fn emit_str_delete(chunks: &mut [Chunk], current: usize, line: u32) {
    use vybe_runtime::opcode::Op;
    let count = chunks[current].alloc_scratch(1);
    let index = chunks[current].alloc_scratch(1);
    let s = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, count, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, index, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, s, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, index, line);
    // Pascal numbers are f64 on the stack, so the 1-based adjustment must be
    // f64 arithmetic — `i32.sub` against an `f64.const` silently misbehaves.
    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_op(Op::F64_SUB, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, count, line);
    chunks[current].emit_string_const("", line);
    vybe_compiler::primitives::strings::emit_splice(chunks, current, line);
}

/// Pascal `Insert(Src; var Dst; Index)` — stack `[src, dst, index]`.
///
/// Same splice with nothing deleted; note the source comes FIRST, which is why
/// this cannot be the same binding as `Delete`.
pub fn emit_str_insert_var(chunks: &mut [Chunk], current: usize, line: u32) {
    use vybe_runtime::opcode::Op;
    let index = chunks[current].alloc_scratch(1);
    let dst = chunks[current].alloc_scratch(1);
    let src = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, index, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, dst, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, src, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, dst, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, index, line);
    // Pascal numbers are f64 on the stack, so the 1-based adjustment must be
    // f64 arithmetic — `i32.sub` against an `f64.const` silently misbehaves.
    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_op(Op::F64_SUB, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, src, line);
    vybe_compiler::primitives::strings::emit_splice(chunks, current, line);
}

// ── Delphi's `Generics.Collections` quirks ───────────────────────────────────
//
// A `TList` IS the shared array — `Add`, `Count`, `Delete` and the rest bind
// straight to `collections.*`. These are the members that do NOT have a shared
// concept behind them, or that need an argument the binding cannot carry:
// `CommonEmit` is a name with no bound value, so `First` cannot simply BE
// `collections.get`. They decompose into the same shared routes here, in the
// language that speaks Delphi, so nothing common learns the word.

/// `L.First` — stack `[arr]` → `[value]`. Delphi's spelling for `L[0]`.
pub fn emit_list_first(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_f64_const(0.0, line);
    collections::emit_get(chunks, current, line);
}

/// `L.Last` — stack `[arr]` → `[value]`. `L[Count - 1]`.
pub fn emit_list_last(chunks: &mut [Chunk], current: usize, line: u32) {
    let arr = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, arr, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_op(Op::F64_SUB, line);
    collections::emit_get(chunks, current, line);
}

/// `L.Exchange(i, j)` — stack `[arr, i, j]` → `[null]`. A swap through the
/// shared get/set; Delphi has a name for it, the store does not.
pub fn emit_list_exchange(chunks: &mut [Chunk], current: usize, line: u32) {
    let j = chunks[current].alloc_scratch(1);
    let i = chunks[current].alloc_scratch(1);
    let arr = chunks[current].alloc_scratch(1);
    let tmp = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, j, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, arr, line);

    // tmp := arr[i]
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, tmp, line);
    // arr[i] := arr[j]
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, j, line);
    collections::emit_get(chunks, current, line);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    // arr[j] := tmp
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, j, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, tmp, line);
    collections::emit_set(chunks, current, line);
}

/// `L.Move(cur, new)` — stack `[arr, cur, new]` → `[null]`. Remove, re-insert.
///
/// The synthesized prelude's version of this was `FItems[0] := FItems[1];
/// FItems[1] := FItems[2]; FItems[2] := value` — a hardcoded three-element
/// shuffle that answered its own test and nothing else.
pub fn emit_list_move(chunks: &mut [Chunk], current: usize, line: u32) {
    let dst = chunks[current].alloc_scratch(1);
    let src = chunks[current].alloc_scratch(1);
    let arr = chunks[current].alloc_scratch(1);
    let val = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, dst, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, src, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, arr, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, arr, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, src, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, val, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, arr, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, src, line);
    collections::emit_remove_at(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, arr, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, dst, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, val, line);
    collections::emit_insert_at(chunks, current, line);
}

/// `L.AddRange(src)` — stack `[arr, src]` → `[null]`. Append at the end.
pub fn emit_list_add_range(chunks: &mut [Chunk], current: usize, line: u32) {
    let src = chunks[current].alloc_scratch(1);
    let arr = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, src, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, arr, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, arr, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, src, line);
    collections::emit_insert_range(chunks, current, line);
}

/// `L.ExtractAt(i)` — stack `[arr, i]` → `[value]`. This IS `list.pop(i)`;
/// the shared `remove_at` discards what it removed, so read it first.
pub fn emit_list_extract_at(chunks: &mut [Chunk], current: usize, line: u32) {
    let idx = chunks[current].alloc_scratch(1);
    let arr = chunks[current].alloc_scratch(1);
    let val = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, arr, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, arr, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, val, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, arr, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx, line);
    collections::emit_remove_at(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, val, line);
}

/// `L.Extract(v)` — stack `[arr, v]` → `[value]`. Remove by VALUE and hand it
/// back; the shared route answers a bool, and the value is already in hand.
pub fn emit_list_extract(chunks: &mut [Chunk], current: usize, line: u32) {
    let val = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, val, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, val, line);
    collections::emit_remove_value(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, val, line);
}

/// `L.TrimExcess` / `L.Capacity := n` — stack `[arr, …]` → `[null]`.
///
/// Capacity is a manual-allocation concept. A growable shared array has no
/// spare capacity to trim, so this is a no-op that keeps the stack honest.
pub fn emit_list_drop_args(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    for _ in 0..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

// ── TDictionary — the shared Map ─────────────────────────────────────────
//
// `TDictionary` is an `ObjectKind::Map`: the SAME store a PHP array and a
// Python dict land on, so a Delphi dictionary handed to either is a thing
// they already understand.
//
// It is NOT the `common:dict.*` family. Those helpers are the older
// Ordinary-object shape and read a `__keys` property for enumeration —
// `dict::emit_set` is a bare `ARRAY_SET` that never appends to it, so on a
// Map `dict.size` answers 0 and `dict.keys` answers `[]`, both silently.
// PHP and Python each reached the same conclusion and route their Map
// members through `ecma:map.*` for the same reason.
//
// Reads and writes stay on `common:dict.get_dynamic`/`set_dynamic`:
// `ARRAY_GET`/`ARRAY_SET` dispatch on `ObjectKind` and are already
// Map-aware, so `D['a']` needs nothing Pascal-specific.

/// One `ecma:map` call. `argc` operands are already on the stack.
fn map_call(chunks: &mut [Chunk], current: usize, func: &str, argc: u8, line: u32) {
    let f = chunks[current].add_import("ecma:map", func);
    chunks[current].emit_call(f, argc, line);
}

/// `TDictionary.Create` — stack `[]` → `[map]`.
pub fn emit_dict_new(chunks: &mut [Chunk], current: usize, line: u32) {
    map_call(chunks, current, "new", 0, line);
}

/// `D.Count` — stack `[map]` → `[i32]`.
pub fn emit_dict_size(chunks: &mut [Chunk], current: usize, line: u32) {
    map_call(chunks, current, "size", 1, line);
}

/// `D.ContainsKey(k)` — stack `[map, key]` → `[bool]`.
pub fn emit_dict_has(chunks: &mut [Chunk], current: usize, line: u32) {
    map_call(chunks, current, "has", 2, line);
}

/// `D.ContainsValue(v)` — stack `[map, value]` → `[bool]`.
pub fn emit_dict_contains_value(chunks: &mut [Chunk], current: usize, line: u32) {
    map_call(chunks, current, "containsValue", 2, line);
}

/// `D.Remove(k)` — stack `[map, key]` → `[null]`.
///
/// Delphi's `Remove` is a procedure; `ecma:map.delete` answers a bool, which
/// would leave the stack one deep in statement position.
pub fn emit_dict_delete(chunks: &mut [Chunk], current: usize, line: u32) {
    map_call(chunks, current, "delete", 2, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

/// `D.Clear` — stack `[map]` → `[null]`.
pub fn emit_dict_clear(chunks: &mut [Chunk], current: usize, line: u32) {
    map_call(chunks, current, "clear", 1, line);
}

/// `D.Keys` / `D.Values` / `D.ToArray` — stack `[map]` → `[array]`.
///
/// `ecma:map.keys` yields an Array Iterator per ECMA-262 §24.1.3.8;
/// `ecma:object.*` answers with a materialized array and is Map-aware, which
/// is the shape `for … in D.Keys` and `TPair` iteration both want.
pub fn emit_dict_enumerate(chunks: &mut [Chunk], current: usize, which: &str, line: u32) {
    let f = chunks[current].add_import("ecma:object", which);
    chunks[current].emit_call(f, 1, line);
}

// ── Exceptions — the SHARED exception model ──────────────────────────────
//
// Pascal used to synthesize `Exception` and ten `E*` subclasses as Pascal
// SOURCE in the walker — the same prelude pattern the collections classes
// used, and with the same consequence: the object carried none of the shared
// stamps, so a Pascal `EDivByZero` could not be caught as `Exception` by PHP
// or Java, and never canonicalised to `ZeroDivisionError`.
//
// `primitives/errors.rs` already models this for every language.
// `emit_exception_new_finalize` coerces the message per ECMA-262 §20.5.1.1,
// sets `message`, and stamps `__type`/`__type_name`/`__exception_type` with
// the CANONICAL name — its own comment says why: "for cross-language catch
// dispatch and introspection compatibility". It also keeps the ORIGINAL
// spelling as `name`, so `EDivByZero` still prints as itself.
// `emit_stamp_exception_ancestors` writes the `__types` MRO that a typed
// `catch`/`on E: ... do` matches against.
//
// Reached from a tree `ctor_call` (see `register_tree` in `lib.rs`), so the
// canonical name is a BOUND ARGUMENT chosen at registration rather than a row
// in a shared table — nothing shared has to learn Pascal's spellings.
//
// NOTE: `primitives/errors.rs` is the one substantial primitive with no
// `common:errors.*` dispatch category (the categories are exactly the module
// names — collections, csv, dict, math, object, reflection, sprintf, strings,
// url, xml). Until it has one, a language reaches it by calling in from its
// own emitter, which is what java and php do too.

/// `E<Something>.Create(msg)` — stack `[msg]` → `[exception]`.
///
/// `canonical` is the shared name this Pascal spelling maps to; `spelling` is
/// what the source called it and what `name` keeps.
pub fn emit_exception_new(chunks: &mut [Chunk], current: usize, spelling: &str, line: u32) {
    let msg = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, msg, line);

    // The shape `emit_exception_new_finalize` expects: [obj, obj, msg].
    class_slots::emit_class_alloc(&mut chunks[current], line);
    chunks[current].emit_dup(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, msg, line);
    vybe_compiler::primitives::errors::emit_exception_new_finalize(
        &mut chunks[current],
        spelling,
        line,
    );
    vybe_compiler::primitives::errors::emit_stamp_exception_ancestors(
        &mut chunks[current],
        spelling,
        line,
    );
}

// ── Variants ────────────────────────────────────────────────────────────────
//
// A `Variant` IS the dynamic value the VM already holds — there is no tag
// beside it and none is needed, because the value knows its own type. So
// `VarType` is a reading of that runtime type, and every `VarIsX` predicate is
// a comparison against the reading rather than a separate probe.
//
// The numbers are `Variants.pas`'s own, because a program compares against the
// SPELLING (`VarType(v) = varInteger`) and the profile declares those constants
// with these same values. Free Pascal, which is the ground truth here, gives a
// string variant `varString` — Delphi's Unicode `varUString` is not what `fpc`
// produces for `v := 'text'`.

/// `varEmpty` — assigned nothing. The VM's `Undefined`.
const VAR_EMPTY: i32 = 0;
/// `varNull` — assigned `Null`, which is a VALUE and not the absence of one.
const VAR_NULL: i32 = 1;
const VAR_INTEGER: i32 = 3;
const VAR_DOUBLE: i32 = 5;
const VAR_BOOLEAN: i32 = 11;
const VAR_STRING: i32 = 256;
/// `varArray` is a FLAG or'd onto the element type, not a type of its own.
const VAR_ARRAY: i32 = 0x2000;

/// Stack `[x]` → `[i32 0|1]`: is the string on top equal to `want`?
fn emit_str_is(chunks: &mut [Chunk], current: usize, want: &str, line: u32) {
    chunks[current].emit_string_const(want, line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
}

/// `VarType(v)` — stack `[v]` → `[code]`.
///
/// `typeof` separates every case except one: it answers `object` for both an
/// array and `Null`, and `undefined` is its own answer, which is exactly the
/// `varEmpty`/`varNull` distinction Delphi draws and `ref.is_null` cannot —
/// that opcode is true for BOTH.
pub fn emit_var_type(chunks: &mut [Chunk], current: usize, line: u32) {
    let v = chunks[current].alloc_scratch(2);
    let t = v + 1;
    chunks[current].emit_op_u16(Op::LOCAL_SET, v, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, v, line);
    host(chunks, current, "ecma:value", "typeof", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, t, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, t, line);
    emit_str_is(chunks, current, "undefined", line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_i32_const(VAR_EMPTY, line);
    chunks[current].emit_else(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, t, line);
    emit_str_is(chunks, current, "boolean", line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_i32_const(VAR_BOOLEAN, line);
    chunks[current].emit_else(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, t, line);
    emit_str_is(chunks, current, "string", line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_i32_const(VAR_STRING, line);
    chunks[current].emit_else(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, t, line);
    emit_str_is(chunks, current, "number", line);
    chunks[current].emit_if_value(line);
    // Delphi stores a whole number as `varInteger` and a fractional one as
    // `varDouble`; the VM holds both as one numeric value, so the DISTINCTION
    // is the value's own integrality.
    chunks[current].emit_op_u16(Op::LOCAL_GET, v, line);
    host(chunks, current, "ecma:number", "isInteger", 1, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_i32_const(VAR_INTEGER, line);
    chunks[current].emit_else(line);
    chunks[current].emit_i32_const(VAR_DOUBLE, line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);

    // `object`: an array, or `Null`.
    chunks[current].emit_op_u16(Op::LOCAL_GET, v, line);
    host(chunks, current, "ecma:array", "isArray", 1, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_i32_const(VAR_ARRAY | VAR_INTEGER, line);
    chunks[current].emit_else(line);
    chunks[current].emit_i32_const(VAR_NULL, line);
    chunks[current].emit_end(line);

    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

/// `VarType(v) = code` — the shape every single-code predicate has.
fn emit_var_type_is(chunks: &mut [Chunk], current: usize, code: i32, line: u32) {
    emit_var_type(chunks, current, line);
    chunks[current].emit_i32_const(code, line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    ops::emit_i32_to_bool(&mut chunks[current], line);
}

/// `VarType(v)` is either of two codes — `VarIsNumeric`, `VarIsOrdinal`.
fn emit_var_type_is_either(chunks: &mut [Chunk], current: usize, a: i32, b: i32, line: u32) {
    let code = chunks[current].alloc_scratch(1);
    emit_var_type(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, code, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, code, line);
    chunks[current].emit_i32_const(a, line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_bool_const(true, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, code, line);
    chunks[current].emit_i32_const(b, line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    ops::emit_i32_to_bool(&mut chunks[current], line);
    chunks[current].emit_end(line);
}

/// `VarIsEmpty` / `VarIsNull` / `VarIsStr` / `VarIsBool` / `VarIsFloat` /
/// `VarIsNumeric` / `VarIsOrdinal` / `VarIsArray` / `VarIsByRef`.
pub fn emit_var_is(chunks: &mut [Chunk], current: usize, which: &str, line: u32) -> bool {
    match which {
        "empty" => emit_var_type_is(chunks, current, VAR_EMPTY, line),
        "null" => emit_var_type_is(chunks, current, VAR_NULL, line),
        "str" => emit_var_type_is(chunks, current, VAR_STRING, line),
        "bool" => emit_var_type_is(chunks, current, VAR_BOOLEAN, line),
        "float" => emit_var_type_is(chunks, current, VAR_DOUBLE, line),
        // Delphi's ordinal types are the integer family, `Boolean` and the
        // enumerations — everything counted, as opposed to measured.
        "ordinal" => emit_var_type_is_either(chunks, current, VAR_INTEGER, VAR_BOOLEAN, line),
        "numeric" => emit_var_type_is_either(chunks, current, VAR_INTEGER, VAR_DOUBLE, line),
        "array" => {
            host(chunks, current, "ecma:array", "isArray", 1, line);
            ops::emit_dyn_to_bool(&mut chunks[current], line);
        }
        // A variant here never wraps a reference to someone else's storage:
        // `VarIsByRef` describes `varByRef`, which only an interop caller can
        // set, and nothing in this runtime produces one.
        "byref" => {
            chunks[current].emit_op(Op::DROP, line);
            chunks[current].emit_bool_const(false, line);
        }
        _ => return false,
    }
    true
}

/// `VarToStr(v)` — stack `[v]` → `[string]`. `Null` and `Unassigned` render as
/// the EMPTY string, which is the one place `VarToStr` differs from `String()`.
pub fn emit_var_to_str(chunks: &mut [Chunk], current: usize, line: u32) {
    let v = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, v, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, v, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_string_const("", line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, v, line);
    host(chunks, current, "ecma:string", "String", 1, line);
    chunks[current].emit_end(line);
}

/// `VarToStrDef(v, default)` — stack `[v, default]` → `[string]`.
pub fn emit_var_to_str_def(chunks: &mut [Chunk], current: usize, line: u32) {
    let dflt = chunks[current].alloc_scratch(2);
    let v = dflt + 1;
    chunks[current].emit_op_u16(Op::LOCAL_SET, dflt, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, v, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, v, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, dflt, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, v, line);
    host(chunks, current, "ecma:string", "String", 1, line);
    chunks[current].emit_end(line);
}

/// `VarTypeToAsString(code)` — the type code's own name, as `Variants.pas`
/// spells it in a message.
pub fn emit_var_type_to_as_string(chunks: &mut [Chunk], current: usize, line: u32) {
    let code = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, code, line);
    for (value, name) in [
        (VAR_EMPTY, "Empty"),
        (VAR_NULL, "Null"),
        (VAR_INTEGER, "Integer"),
        (VAR_DOUBLE, "Double"),
        (VAR_BOOLEAN, "Boolean"),
        (VAR_STRING, "String"),
    ] {
        chunks[current].emit_op_u16(Op::LOCAL_GET, code, line);
        chunks[current].emit_i32_const(value, line);
        ops::emit_dyn_eq(&mut chunks[current], line);
        ops::emit_dyn_to_bool(&mut chunks[current], line);
        chunks[current].emit_if_value(line);
        chunks[current].emit_string_const(name, line);
        chunks[current].emit_else(line);
    }
    chunks[current].emit_string_const("Unknown", line);
    for _ in 0..6 {
        chunks[current].emit_end(line);
    }
}

/// `VarSameValue(a, b)` — variant equality, which compares VALUES across types
/// (`1 = '1'`) exactly as the variant `=` operator does.
pub fn emit_var_same_value(chunks: &mut [Chunk], current: usize, line: u32) {
    ops::emit_dyn_eq(&mut chunks[current], line);
    ops::emit_i32_to_bool(&mut chunks[current], line);
}

/// `VarAsType(v, code)` — stack `[v, code]` → `[coerced]`.
///
/// A coercion that cannot be performed raises `EVariantError`; that is the
/// documented contract, and `VarAsType('NotANumber', varInteger)` is the case
/// a program catches.
pub fn emit_var_as_type(chunks: &mut [Chunk], current: usize, line: u32) {
    let code = chunks[current].alloc_scratch(3);
    let v = code + 1;
    let out = code + 2;
    chunks[current].emit_op_u16(Op::LOCAL_SET, code, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, v, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, code, line);
    chunks[current].emit_i32_const(VAR_STRING, line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, v, line);
    host(chunks, current, "ecma:string", "String", 1, line);
    chunks[current].emit_else(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, code, line);
    chunks[current].emit_i32_const(VAR_BOOLEAN, line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, v, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_else(line);

    // Every remaining code is numeric. `Number('12x')` is NaN, and NaN is
    // precisely "this variant does not convert".
    chunks[current].emit_op_u16(Op::LOCAL_GET, v, line);
    host(chunks, current, "ecma:number", "Number", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
    ops::emit_dyn_ne(&mut chunks[current], line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_string_const("Could not convert variant to the required type", line);
    emit_exception_new(chunks, current, "EVariantError", line);
    vybe_compiler::primitives::errors::emit_throw(&mut chunks[current], line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, code, line);
    chunks[current].emit_i32_const(VAR_INTEGER, line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
    vybe_compiler::primitives::math::emit_trunc(&mut chunks[current], line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
    chunks[current].emit_end(line);

    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}
