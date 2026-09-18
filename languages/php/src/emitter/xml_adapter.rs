//! PHP XML adapter — `SimpleXML` + DOM serialization, as inline-emit
//! functions composing the ECMA `web:dom-parser` host surface (DOMParser /
//! Document / Element). No PHP-specific host fns; everything is bytecode
//! emission over the spec-conformant DOM host, mirroring the other
//! `emitter/php/*_adapter.rs` modules.
//!
//! PHP `DOMDocument` method calls are routed straight to `web:dom-parser`
//! via the walker (`$node->createElement(...)` → `__dom_createElement(node,
//! ...)`), so those need no emit here. This module carries the parts that
//! need real composition: `simplexml_load_string` (parse → root element)
//! and the SimpleXML value shape.

use std::sync::Arc;
use vybe_compiler::primitives::class_slots::{
    self, ClassSlot, Dest, ObjSource, PlainNames, ValueSource,
};
use vybe_runtime::opcode::Op;
use vybe_runtime::{Chunk, Value};

fn call_import(
    chunks: &mut [Chunk],
    current: usize,
    module: &str,
    name: &str,
    argc: u8,
    line: u32,
) {
    let idx = chunks[current].add_import(module.to_string(), name.to_string());
    chunks[current].emit_call(idx, argc, line);
}

fn struct_get_key(chunk: &mut Chunk, key: &ClassSlot, line: u32) {
    let slot = class_slots::resolve(key, &PlainNames);
    class_slots::emit_class_get(chunk, ObjSource::Stack, &slot, Dest::Stack, line);
}

fn struct_set_key(chunk: &mut Chunk, key: &ClassSlot, line: u32) {
    let slot = class_slots::resolve(key, &PlainNames);
    class_slots::emit_class_set(chunk, ObjSource::Stack, &slot, ValueSource::Stack, line);
}

fn lset(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}

fn lget(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn str_concat(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("wasm:js-string", "concat");
    chunk.emit_call(idx, 2, line);
}

fn get_field(chunk: &mut Chunk, obj_slot: u16, field: &str, line: u32) {
    lget(chunk, obj_slot, line);
    struct_get_key(chunk, &ClassSlot::internal(field), line);
}

fn set_field_from_stack(chunk: &mut Chunk, obj_slot: u16, field: &str, line: u32) {
    let value_slot = chunk.alloc_scratch(1);
    lset(chunk, value_slot, line);
    lget(chunk, obj_slot, line);
    lget(chunk, value_slot, line);
    struct_set_key(chunk, &ClassSlot::internal(field), line);
}

fn set_field_str(chunk: &mut Chunk, obj_slot: u16, field: &str, value: &str, line: u32) {
    lget(chunk, obj_slot, line);
    chunk.emit_string_const(value, line);
    struct_set_key(chunk, &ClassSlot::internal(field), line);
}

fn set_field_bool(chunk: &mut Chunk, obj_slot: u16, field: &str, value: bool, line: u32) {
    lget(chunk, obj_slot, line);
    chunk.emit_bool_const(value, line);
    struct_set_key(chunk, &ClassSlot::internal(field), line);
}

fn append_slot_to_buf(chunk: &mut Chunk, obj_slot: u16, value_slot: u16, line: u32) {
    let out_slot = chunk.alloc_scratch(1);
    get_field(chunk, obj_slot, "__buf", line);
    lget(chunk, value_slot, line);
    str_concat(chunk, line);
    lset(chunk, out_slot, line);
    lget(chunk, obj_slot, line);
    lget(chunk, out_slot, line);
    struct_set_key(chunk, &ClassSlot::internal("__buf"), line);
}

fn append_str_to_buf(chunk: &mut Chunk, obj_slot: u16, value: &str, line: u32) {
    let value_slot = chunk.alloc_scratch(1);
    chunk.emit_string_const(value, line);
    lset(chunk, value_slot, line);
    append_slot_to_buf(chunk, obj_slot, value_slot, line);
}

fn append_slot_to_attr(chunk: &mut Chunk, obj_slot: u16, value_slot: u16, line: u32) {
    let out_slot = chunk.alloc_scratch(1);
    get_field(chunk, obj_slot, "__attr_buf", line);
    lget(chunk, value_slot, line);
    str_concat(chunk, line);
    lset(chunk, out_slot, line);
    lget(chunk, obj_slot, line);
    lget(chunk, out_slot, line);
    struct_set_key(chunk, &ClassSlot::internal("__attr_buf"), line);
}

fn close_open_tag_if_needed(chunk: &mut Chunk, obj_slot: u16, line: u32) {
    let open_slot = chunk.alloc_scratch(1);
    let closed_slot = chunk.alloc_scratch(1);
    get_field(chunk, obj_slot, "__open", line);
    lset(chunk, open_slot, line);
    lget(chunk, open_slot, line);
    chunk.emit_string_const("", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_if_value(line);
    lget(chunk, open_slot, line);
    chunk.emit_string_const(">", line);
    str_concat(chunk, line);
    lset(chunk, closed_slot, line);
    append_slot_to_buf(chunk, obj_slot, closed_slot, line);
    set_field_str(chunk, obj_slot, "__open", "", line);
    chunk.emit_end(line);
}

fn xmlwriter_push_name(chunk: &mut Chunk, obj_slot: u16, name_slot: u16, line: u32) {
    get_field(chunk, obj_slot, "__stack", line);
    lget(chunk, name_slot, line);
    let push = chunk.add_import("ecma:array", "push");
    chunk.emit_call(push, 2, line);
    chunk.emit_op(Op::DROP, line);
}

fn xmlwriter_set_open_from_name(chunk: &mut Chunk, obj_slot: u16, name_slot: u16, line: u32) {
    chunk.emit_string_const("<", line);
    lget(chunk, name_slot, line);
    str_concat(chunk, line);
    set_field_from_stack(chunk, obj_slot, "__open", line);
}

fn xmlwriter_append_element(
    chunk: &mut Chunk,
    obj_slot: u16,
    name_slot: u16,
    value_slot: u16,
    line: u32,
) {
    get_field(chunk, obj_slot, "__indent", line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    append_str_to_buf(chunk, obj_slot, "\n  ", line);
    chunk.emit_end(line);

    append_str_to_buf(chunk, obj_slot, "<", line);
    append_slot_to_buf(chunk, obj_slot, name_slot, line);
    append_str_to_buf(chunk, obj_slot, ">", line);
    append_slot_to_buf(chunk, obj_slot, value_slot, line);
    append_str_to_buf(chunk, obj_slot, "</", line);
    append_slot_to_buf(chunk, obj_slot, name_slot, line);
    append_str_to_buf(chunk, obj_slot, ">", line);
}

/// PHP `$doc->saveXML($node?)` — serialize the node via the ECMA
/// `XMLSerializer` host and append the trailing newline PHP emits (which
/// `serializeToString` does not). The node is already on the stack.
pub fn emit_dom_save_xml(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    call_import(
        chunks,
        current,
        "web:dom-parser",
        "serializeToString",
        1,
        line,
    );
    let chunk = &mut chunks[current];
    chunk.emit_string_const("\n", line);
    let idx = chunk.add_import("wasm:js-string", "concat");
    chunk.emit_call(idx, 2, line);
}

/// PHP `simplexml_load_string($xml [, $class, $opts, $ns, $prefix])` — parse
/// the XML string through the ECMA DOM host and return the document's root
/// element (`documentElement`). PHP's `false`-on-empty/error is handled by
/// the host `parse` returning a document with a null `documentElement`,
/// which the SimpleXML access layer treats as absent.
pub fn emit_simplexml_load_string(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    // Drop the optional class/options/namespace/prefix arguments — the MVP
    // uses the default SimpleXMLElement shape.
    {
        let chunk = &mut chunks[current];
        for _ in 1..argc {
            chunk.emit_op(Op::DROP, line);
        }
    }
    // parse(xml) → Document
    call_import(chunks, current, "web:dom-parser", "parse", 1, line);
    let chunk = &mut chunks[current];
    let doc_slot = chunk.alloc_scratch(2);
    let root_slot = doc_slot + 1;
    lset(chunk, doc_slot, line);
    // → documentElement (the SimpleXML root)
    lget(chunk, doc_slot, line);
    struct_get_key(chunk, &ClassSlot::internal("documentElement"), line);
    chunk.emit_op_u16(Op::LOCAL_TEE, root_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if_value(line);
    chunk.emit_bool_const(false, line);
    chunk.emit_else(line);
    lget(chunk, root_slot, line);
    struct_get_key(chunk, &ClassSlot::internal("nodeName"), line);
    chunk.emit_string_const("parsererror", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    chunk.emit_bool_const(false, line);
    chunk.emit_else(line);
    lget(chunk, root_slot, line);
    chunk.emit_end(line);
    chunk.emit_end(line);
}

/// `new DOMXPath($doc)` adapter object. The walker resolves methods back to
/// the `document` field, while `$xpath->document` remains observable.
pub fn emit_domxpath_new(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    if argc == 0 {
        chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    } else {
        for _ in 1..argc {
            chunk.emit_op(Op::DROP, line);
        }
    }
    let doc_slot = chunk.alloc_scratch(2);
    let obj_slot = doc_slot + 1;
    lset(chunk, doc_slot, line);
    class_slots::emit_class_alloc(chunk, line);
    lset(chunk, obj_slot, line);
    lget(chunk, obj_slot, line);
    chunk.emit_string_const("DOMXPath", line);
    struct_set_key(chunk, &ClassSlot::TypeIdentity, line);
    lget(chunk, obj_slot, line);
    lget(chunk, doc_slot, line);
    struct_set_key(chunk, &ClassSlot::internal("document"), line);
    lget(chunk, obj_slot, line);
}

/// DOMAttr value object used by `DOMDocument::createAttribute`.
pub fn emit_dom_attr_new(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    if argc == 0 {
        chunk.emit_string_const("", line);
    } else {
        for _ in 1..argc {
            chunk.emit_op(Op::DROP, line);
        }
    }
    let name_slot = chunk.alloc_scratch(2);
    let obj_slot = name_slot + 1;
    lset(chunk, name_slot, line);
    class_slots::emit_class_alloc(chunk, line);
    lset(chunk, obj_slot, line);
    lget(chunk, obj_slot, line);
    chunk.emit_string_const("DOMAttr", line);
    struct_set_key(chunk, &ClassSlot::TypeIdentity, line);
    lget(chunk, obj_slot, line);
    lget(chunk, name_slot, line);
    struct_set_key(chunk, &ClassSlot::internal("name"), line);
    lget(chunk, obj_slot, line);
    chunk.emit_string_const("", line);
    struct_set_key(chunk, &ClassSlot::internal("value"), line);
    lget(chunk, obj_slot, line);
}

/// SimpleXML property read: first descendant with the requested tag, as text.
pub fn emit_simplexml_child_text(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    call_import(
        chunks,
        current,
        "web:dom-parser",
        "getElementsByTagName",
        2,
        line,
    );
    let chunk = &mut chunks[current];
    chunk.emit_i32_const(0, line);
    let get = chunk.add_import("ecma:array", "get");
    chunk.emit_call(get, 2, line);
    struct_get_key(chunk, &ClassSlot::internal("textContent"), line);
}

pub fn emit_dom_node_item(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    call_import(chunks, current, "ecma:array", "get", 2, line);
}

pub fn emit_php_xmlwriter_new(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    class_slots::emit_class_alloc(chunk, line);
    let obj_slot = chunk.alloc_scratch(1);
    lset(chunk, obj_slot, line);
    lget(chunk, obj_slot, line);
    chunk.emit_string_const("XMLWriter", line);
    struct_set_key(chunk, &ClassSlot::TypeIdentity, line);
    set_field_str(chunk, obj_slot, "__buf", "", line);
    set_field_str(chunk, obj_slot, "__open", "", line);
    lget(chunk, obj_slot, line);
    chunk.emit_array_new_fixed(0, 0, line);
    struct_set_key(chunk, &ClassSlot::internal("__stack"), line);
    set_field_bool(chunk, obj_slot, "__indent", false, line);
    set_field_str(chunk, obj_slot, "__indent_string", "  ", line);
    set_field_bool(chunk, obj_slot, "__in_attr", false, line);
    set_field_str(chunk, obj_slot, "__attr_name", "", line);
    set_field_str(chunk, obj_slot, "__attr_buf", "", line);
    lget(chunk, obj_slot, line);
}

pub fn emit_php_xmlwriter_open_memory(
    chunks: &mut Vec<Chunk>,
    current: usize,
    _argc: u8,
    line: u32,
) {
    let chunk = &mut chunks[current];
    let obj_slot = chunk.alloc_scratch(1);
    lset(chunk, obj_slot, line);
    set_field_str(chunk, obj_slot, "__buf", "", line);
    set_field_str(chunk, obj_slot, "__open", "", line);
    lget(chunk, obj_slot, line);
    chunk.emit_array_new_fixed(0, 0, line);
    struct_set_key(chunk, &ClassSlot::internal("__stack"), line);
    chunk.emit_bool_const(true, line);
}

pub fn emit_php_xmlwriter_start_document(
    chunks: &mut Vec<Chunk>,
    current: usize,
    argc: u8,
    line: u32,
) {
    let chunk = &mut chunks[current];
    let enc_slot = chunk.alloc_scratch(3);
    let version_slot = enc_slot + 1;
    let obj_slot = enc_slot + 2;
    if argc >= 3 {
        lset(chunk, enc_slot, line);
    } else {
        chunk.emit_string_const("UTF-8", line);
        lset(chunk, enc_slot, line);
    }
    if argc >= 2 {
        lset(chunk, version_slot, line);
    } else {
        chunk.emit_string_const("1.0", line);
        lset(chunk, version_slot, line);
    }
    lset(chunk, obj_slot, line);
    append_str_to_buf(chunk, obj_slot, "<?xml version=\"", line);
    append_slot_to_buf(chunk, obj_slot, version_slot, line);
    append_str_to_buf(chunk, obj_slot, "\" encoding=\"", line);
    append_slot_to_buf(chunk, obj_slot, enc_slot, line);
    append_str_to_buf(chunk, obj_slot, "\"?>\n", line);
    chunk.emit_bool_const(true, line);
}

pub fn emit_php_xmlwriter_start_element(
    chunks: &mut Vec<Chunk>,
    current: usize,
    argc: u8,
    line: u32,
) {
    let chunk = &mut chunks[current];
    let name_slot = chunk.alloc_scratch(2);
    let obj_slot = name_slot + 1;
    if argc >= 2 {
        lset(chunk, name_slot, line);
    } else {
        chunk.emit_string_const("", line);
        lset(chunk, name_slot, line);
    }
    lset(chunk, obj_slot, line);
    close_open_tag_if_needed(chunk, obj_slot, line);
    xmlwriter_push_name(chunk, obj_slot, name_slot, line);
    xmlwriter_set_open_from_name(chunk, obj_slot, name_slot, line);
    chunk.emit_bool_const(true, line);
}

pub fn emit_php_xmlwriter_start_element_ns(
    chunks: &mut Vec<Chunk>,
    current: usize,
    argc: u8,
    line: u32,
) {
    let chunk = &mut chunks[current];
    let ns_slot = chunk.alloc_scratch(5);
    let local_slot = ns_slot + 1;
    let prefix_slot = ns_slot + 2;
    let obj_slot = ns_slot + 3;
    let name_slot = ns_slot + 4;
    if argc >= 4 {
        lset(chunk, ns_slot, line);
    } else {
        chunk.emit_string_const("", line);
        lset(chunk, ns_slot, line);
    }
    if argc >= 3 {
        lset(chunk, local_slot, line);
    } else {
        chunk.emit_string_const("", line);
        lset(chunk, local_slot, line);
    }
    if argc >= 2 {
        lset(chunk, prefix_slot, line);
    } else {
        chunk.emit_string_const("", line);
        lset(chunk, prefix_slot, line);
    }
    lset(chunk, obj_slot, line);
    lget(chunk, prefix_slot, line);
    chunk.emit_string_const(":", line);
    str_concat(chunk, line);
    lget(chunk, local_slot, line);
    str_concat(chunk, line);
    lset(chunk, name_slot, line);
    close_open_tag_if_needed(chunk, obj_slot, line);
    xmlwriter_push_name(chunk, obj_slot, name_slot, line);
    chunk.emit_string_const("<", line);
    lget(chunk, name_slot, line);
    str_concat(chunk, line);
    chunk.emit_string_const(" xmlns:", line);
    str_concat(chunk, line);
    lget(chunk, prefix_slot, line);
    str_concat(chunk, line);
    chunk.emit_string_const("=\"", line);
    str_concat(chunk, line);
    lget(chunk, ns_slot, line);
    str_concat(chunk, line);
    chunk.emit_string_const("\"", line);
    str_concat(chunk, line);
    set_field_from_stack(chunk, obj_slot, "__open", line);
    chunk.emit_bool_const(true, line);
}

pub fn emit_php_xmlwriter_write_attribute(
    chunks: &mut Vec<Chunk>,
    current: usize,
    _argc: u8,
    line: u32,
) {
    let chunk = &mut chunks[current];
    let value_slot = chunk.alloc_scratch(4);
    let name_slot = value_slot + 1;
    let obj_slot = value_slot + 2;
    let open_slot = value_slot + 3;
    lset(chunk, value_slot, line);
    lset(chunk, name_slot, line);
    lset(chunk, obj_slot, line);
    get_field(chunk, obj_slot, "__open", line);
    chunk.emit_string_const(" ", line);
    str_concat(chunk, line);
    lget(chunk, name_slot, line);
    str_concat(chunk, line);
    chunk.emit_string_const("=\"", line);
    str_concat(chunk, line);
    lget(chunk, value_slot, line);
    str_concat(chunk, line);
    chunk.emit_string_const("\"", line);
    str_concat(chunk, line);
    lset(chunk, open_slot, line);
    lget(chunk, obj_slot, line);
    lget(chunk, open_slot, line);
    struct_set_key(chunk, &ClassSlot::internal("__open"), line);
    chunk.emit_bool_const(true, line);
}

pub fn emit_php_xmlwriter_start_attribute(
    chunks: &mut Vec<Chunk>,
    current: usize,
    _argc: u8,
    line: u32,
) {
    let chunk = &mut chunks[current];
    let name_slot = chunk.alloc_scratch(2);
    let obj_slot = name_slot + 1;
    lset(chunk, name_slot, line);
    lset(chunk, obj_slot, line);
    lget(chunk, obj_slot, line);
    lget(chunk, name_slot, line);
    struct_set_key(chunk, &ClassSlot::internal("__attr_name"), line);
    set_field_str(chunk, obj_slot, "__attr_buf", "", line);
    set_field_bool(chunk, obj_slot, "__in_attr", true, line);
    chunk.emit_bool_const(true, line);
}

pub fn emit_php_xmlwriter_end_attribute(
    chunks: &mut Vec<Chunk>,
    current: usize,
    _argc: u8,
    line: u32,
) {
    let chunk = &mut chunks[current];
    let obj_slot = chunk.alloc_scratch(3);
    let name_slot = obj_slot + 1;
    let value_slot = obj_slot + 2;
    lset(chunk, obj_slot, line);
    get_field(chunk, obj_slot, "__attr_name", line);
    lset(chunk, name_slot, line);
    get_field(chunk, obj_slot, "__attr_buf", line);
    lset(chunk, value_slot, line);
    lget(chunk, obj_slot, line);
    lget(chunk, name_slot, line);
    lget(chunk, value_slot, line);
    let _ = chunk;
    emit_php_xmlwriter_write_attribute(chunks, current, 3, line);
    let chunk = &mut chunks[current];
    chunk.emit_op(Op::DROP, line);
    set_field_bool(chunk, obj_slot, "__in_attr", false, line);
    chunk.emit_bool_const(true, line);
}

pub fn emit_php_xmlwriter_text(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let text_slot = chunk.alloc_scratch(2);
    let obj_slot = text_slot + 1;
    lset(chunk, text_slot, line);
    lset(chunk, obj_slot, line);
    get_field(chunk, obj_slot, "__in_attr", line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    append_slot_to_attr(chunk, obj_slot, text_slot, line);
    chunk.emit_else(line);
    close_open_tag_if_needed(chunk, obj_slot, line);
    append_slot_to_buf(chunk, obj_slot, text_slot, line);
    chunk.emit_end(line);
    chunk.emit_bool_const(true, line);
}

pub fn emit_php_xmlwriter_write_element(
    chunks: &mut Vec<Chunk>,
    current: usize,
    _argc: u8,
    line: u32,
) {
    let chunk = &mut chunks[current];
    let value_slot = chunk.alloc_scratch(3);
    let name_slot = value_slot + 1;
    let obj_slot = value_slot + 2;
    lset(chunk, value_slot, line);
    lset(chunk, name_slot, line);
    lset(chunk, obj_slot, line);
    close_open_tag_if_needed(chunk, obj_slot, line);
    xmlwriter_append_element(chunk, obj_slot, name_slot, value_slot, line);
    chunk.emit_bool_const(true, line);
}

pub fn emit_php_xmlwriter_end_element(
    chunks: &mut Vec<Chunk>,
    current: usize,
    _argc: u8,
    line: u32,
) {
    let chunk = &mut chunks[current];
    let obj_slot = chunk.alloc_scratch(2);
    let name_slot = obj_slot + 1;
    lset(chunk, obj_slot, line);
    close_open_tag_if_needed(chunk, obj_slot, line);
    get_field(chunk, obj_slot, "__stack", line);
    let pop = chunk.add_import("ecma:array", "pop");
    chunk.emit_call(pop, 1, line);
    lset(chunk, name_slot, line);
    append_str_to_buf(chunk, obj_slot, "</", line);
    append_slot_to_buf(chunk, obj_slot, name_slot, line);
    append_str_to_buf(chunk, obj_slot, ">", line);
    chunk.emit_bool_const(true, line);
}

pub fn emit_php_xmlwriter_output_memory(
    chunks: &mut Vec<Chunk>,
    current: usize,
    _argc: u8,
    line: u32,
) {
    let chunk = &mut chunks[current];
    let obj_slot = chunk.alloc_scratch(1);
    lset(chunk, obj_slot, line);
    close_open_tag_if_needed(chunk, obj_slot, line);
    get_field(chunk, obj_slot, "__buf", line);
}

pub fn emit_php_xmlwriter_flush(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let obj_slot = chunk.alloc_scratch(2);
    let out_slot = obj_slot + 1;
    lset(chunk, obj_slot, line);
    close_open_tag_if_needed(chunk, obj_slot, line);
    get_field(chunk, obj_slot, "__buf", line);
    lset(chunk, out_slot, line);
    set_field_str(chunk, obj_slot, "__buf", "", line);
    lget(chunk, out_slot, line);
}

pub fn emit_php_xmlwriter_end_document(
    chunks: &mut Vec<Chunk>,
    current: usize,
    _argc: u8,
    line: u32,
) {
    let chunk = &mut chunks[current];
    let obj_slot = chunk.alloc_scratch(1);
    lset(chunk, obj_slot, line);
    close_open_tag_if_needed(chunk, obj_slot, line);
    chunk.emit_bool_const(true, line);
}

pub fn emit_php_xmlwriter_set_indent(
    chunks: &mut Vec<Chunk>,
    current: usize,
    _argc: u8,
    line: u32,
) {
    let chunk = &mut chunks[current];
    let value_slot = chunk.alloc_scratch(2);
    let obj_slot = value_slot + 1;
    lset(chunk, value_slot, line);
    lset(chunk, obj_slot, line);
    lget(chunk, obj_slot, line);
    lget(chunk, value_slot, line);
    struct_set_key(chunk, &ClassSlot::internal("__indent"), line);
    chunk.emit_bool_const(true, line);
}

pub fn emit_php_xmlwriter_set_indent_string(
    chunks: &mut Vec<Chunk>,
    current: usize,
    _argc: u8,
    line: u32,
) {
    let chunk = &mut chunks[current];
    let value_slot = chunk.alloc_scratch(2);
    let obj_slot = value_slot + 1;
    lset(chunk, value_slot, line);
    lset(chunk, obj_slot, line);
    lget(chunk, obj_slot, line);
    lget(chunk, value_slot, line);
    struct_set_key(chunk, &ClassSlot::internal("__indent_string"), line);
    chunk.emit_bool_const(true, line);
}

pub fn emit_php_xmlwriter_write_comment(
    chunks: &mut Vec<Chunk>,
    current: usize,
    _argc: u8,
    line: u32,
) {
    let chunk = &mut chunks[current];
    let text_slot = chunk.alloc_scratch(2);
    let obj_slot = text_slot + 1;
    lset(chunk, text_slot, line);
    lset(chunk, obj_slot, line);
    close_open_tag_if_needed(chunk, obj_slot, line);
    append_str_to_buf(chunk, obj_slot, "<!--", line);
    append_slot_to_buf(chunk, obj_slot, text_slot, line);
    append_str_to_buf(chunk, obj_slot, "-->", line);
    chunk.emit_bool_const(true, line);
}

pub fn emit_php_xmlwriter_write_cdata(
    chunks: &mut Vec<Chunk>,
    current: usize,
    _argc: u8,
    line: u32,
) {
    let chunk = &mut chunks[current];
    let text_slot = chunk.alloc_scratch(2);
    let obj_slot = text_slot + 1;
    lset(chunk, text_slot, line);
    lset(chunk, obj_slot, line);
    close_open_tag_if_needed(chunk, obj_slot, line);
    append_str_to_buf(chunk, obj_slot, "<![CDATA[", line);
    append_slot_to_buf(chunk, obj_slot, text_slot, line);
    append_str_to_buf(chunk, obj_slot, "]]>", line);
    chunk.emit_bool_const(true, line);
}

pub fn emit_php_xmlwriter_write_pi(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let data_slot = chunk.alloc_scratch(3);
    let target_slot = data_slot + 1;
    let obj_slot = data_slot + 2;
    if argc >= 3 {
        lset(chunk, data_slot, line);
    } else {
        chunk.emit_string_const("", line);
        lset(chunk, data_slot, line);
    }
    lset(chunk, target_slot, line);
    lset(chunk, obj_slot, line);
    close_open_tag_if_needed(chunk, obj_slot, line);
    append_str_to_buf(chunk, obj_slot, "<?", line);
    append_slot_to_buf(chunk, obj_slot, target_slot, line);
    append_str_to_buf(chunk, obj_slot, " ", line);
    append_slot_to_buf(chunk, obj_slot, data_slot, line);
    append_str_to_buf(chunk, obj_slot, "?>", line);
    chunk.emit_bool_const(true, line);
}

pub fn emit_php_xmlwriter_start_dtd(
    chunks: &mut Vec<Chunk>,
    current: usize,
    _argc: u8,
    line: u32,
) {
    let chunk = &mut chunks[current];
    let name_slot = chunk.alloc_scratch(2);
    let obj_slot = name_slot + 1;
    lset(chunk, name_slot, line);
    lset(chunk, obj_slot, line);
    close_open_tag_if_needed(chunk, obj_slot, line);
    append_str_to_buf(chunk, obj_slot, "<!DOCTYPE ", line);
    append_slot_to_buf(chunk, obj_slot, name_slot, line);
    chunk.emit_bool_const(true, line);
}

pub fn emit_php_xmlwriter_end_dtd(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let obj_slot = chunk.alloc_scratch(1);
    lset(chunk, obj_slot, line);
    append_str_to_buf(chunk, obj_slot, ">", line);
    chunk.emit_bool_const(true, line);
}
