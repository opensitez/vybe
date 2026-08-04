//! Pascal runtime-surface helpers routed via `common:pascal.*`.

use vybe_compiler::primitives::collections;
use vybe_runtime::Chunk;
use vybe_runtime::Op;
use vybe_runtime::Value;

pub fn emit_helper(name: &str, chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) -> bool {
    if name == "pascal.tostring" {
        let to_str = chunks[0].add_import("ecma:string", "String");
        chunks[current].emit_op_u16(Op::CALL_IMPORT, to_str, line);
        chunks[current].emit(argc, line);
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
        let idx = chunks[0].add_import("ecma:set", "size");
        chunks[current].emit_op_u16(Op::CALL_IMPORT, idx, line);
        chunks[current].emit(argc, line);
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
        _ => return false };
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

fn emit_xml_document_new(chunks: &mut [Chunk], current: usize, line: u32) {
    let doc_slot = chunks[current].alloc_scratch(1);
    emit_object_new(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, doc_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, doc_slot, line);
    chunks[current].emit_string_const("__dom", line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunks[current].emit_string_const("", line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    let create = chunks[current].add_import("web:dom-parser", "createDocument");
    chunks[current].emit_call(create, 3, line);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, doc_slot, line);
}

fn emit_object_new(chunks: &mut [Chunk], current: usize, line: u32) {
    let idx = chunks[current].add_import("ecma:object", "new");
    chunks[current].emit_call(idx, 0, line);
}

fn emit_xml_load_from_xml(chunks: &mut [Chunk], current: usize, line: u32) {
    let xml_slot = chunks[current].alloc_scratch(2);
    let doc_slot = xml_slot + 1;
    chunks[current].emit_op_u16(Op::LOCAL_SET, xml_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, doc_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, doc_slot, line);
    chunks[current].emit_string_const("__dom", line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, xml_slot, line);
    vybe_compiler::primitives::xml::emit_parse(chunks, current, 1, line);
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
    chunks[current].emit_string_const("__dom", line);
    collections::emit_get(chunks, current, line);
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

fn emit_xml_child_node(chunks: &mut [Chunk], current: usize, line: u32) {
    let key_slot = chunks[current].alloc_scratch(2);
    let parent_slot = key_slot + 1;
    chunks[current].emit_op_u16(Op::LOCAL_SET, key_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, parent_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, key_slot, line);
    let str_test = chunks[current].add_import("wasm:js-string", "test");
    chunks[current].emit_call(str_test, 1, line);
    chunks[current].emit_if_value(line);
    {
        chunks[current].emit_op_u16(Op::LOCAL_GET, parent_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, key_slot, line);
        vybe_compiler::primitives::xml::emit_elements(chunks, current, 2, line);
        chunks[current].emit_i32_const(0, line);
        collections::emit_get(chunks, current, line);
    }
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, parent_slot, line);
    chunks[current].emit_string_const("childNodes", line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key_slot, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_end(line);
}

fn emit_xml_add_child(chunks: &mut [Chunk], current: usize, line: u32) {
    let name_slot = chunks[current].alloc_scratch(4);
    let parent_slot = name_slot + 1;
    let child_slot = name_slot + 2;
    let dom_slot = name_slot + 3;
    chunks[current].emit_op_u16(Op::LOCAL_SET, name_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, parent_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, parent_slot, line);
    chunks[current].emit_string_const("__dom", line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, dom_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, dom_slot, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, parent_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, dom_slot, line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, name_slot, line);
    let create = chunks[current].add_import("web:dom-parser", "createElement");
    chunks[current].emit_call(create, 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, child_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, dom_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, child_slot, line);
    let append = chunks[current].add_import("web:dom-parser", "appendChild");
    chunks[current].emit_call(append, 2, line);
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
    let stream = chunks[current].alloc_scratch(1);
    let mut slots = Vec::with_capacity(argc as usize);
    for _ in 0..argc {
        let s = chunks[current].alloc_scratch(1);
        chunks[current].emit_op_u16(Op::LOCAL_SET, s, line);
        slots.push(s);
    }
    slots.reverse();

    let stdout_idx = chunks[current].add_import("wasi:cli/stdout", "get-stdout");
    let write_idx = chunks[current].add_import(
        "wasi:io/streams",
        "[method]output-stream.blocking-write-and-flush",
    );
    chunks[current].emit_call(stdout_idx, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, stream, line);

    // Pascal concatenates its arguments with NO separator.
    for s in &slots {
        chunks[current].emit_op_u16(Op::LOCAL_GET, stream, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, *s, line);
        vybe_compiler::primitives::strings::emit_to_string(&mut chunks[current], line);
        chunks[current].emit_call(write_idx, 2, line);
        chunks[current].emit_op(Op::DROP, line);
    }
    if newline {
        chunks[current].emit_op_u16(Op::LOCAL_GET, stream, line);
        chunks[current].emit_string_const("\n", line);
        chunks[current].emit_call(write_idx, 2, line);
        chunks[current].emit_op(Op::DROP, line);
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

/// `BoolToStr(b)` → `true`/`false`; `BoolToStr(b, True)` → `True`/`False`.
pub fn emit_bool_to_str(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc >= 2 {
        chunks[current].emit_op(vybe_runtime::opcode::Op::DROP, line);
    }
    let (t, f) = if argc >= 2 { ("True", "False") } else { ("true", "false") };
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
