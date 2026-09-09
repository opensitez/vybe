use vybe_compiler::primitives::{collections, convert, loops, ops};
use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

const ATTR_KIND: &str = "XAttribute";
const KIND_KEY: &str = "__dotnet_xml_kind";
const NAME_KEY: &str = "name";
const VALUE_KEY: &str = "value";
const MS_BUF: &str = "__ms_buf";
const MS_POS: &str = "__ms_pos";
const MS_LEN: &str = "__ms_len";

fn call_import(
    chunks: &mut [Chunk],
    current: usize,
    module: &str,
    name: &str,
    argc: u8,
    line: u32,
) {
    let idx = chunks[current].add_import(module, name);
    chunks[current].emit_call(idx, argc, line);
}

fn set_field(
    chunks: &mut [Chunk],
    current: usize,
    obj_slot: u16,
    key: &str,
    value_slot: u16,
    line: u32,
) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    chunks[current].emit_string_const(key, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
}

fn set_xml_children_fields(
    chunks: &mut [Chunk],
    current: usize,
    obj_slot: u16,
    value_slot: u16,
    line: u32,
) {
    set_field(chunks, current, obj_slot, "childNodes", value_slot, line);
    set_field(chunks, current, obj_slot, "children", value_slot, line);
}

fn set_field_const_str(
    chunks: &mut [Chunk],
    current: usize,
    obj_slot: u16,
    key: &str,
    value: &str,
    line: u32,
) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    chunks[current].emit_string_const(key, line);
    chunks[current].emit_string_const(value, line);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
}

fn set_field_i32(
    chunks: &mut [Chunk],
    current: usize,
    obj_slot: u16,
    key: &str,
    value: i32,
    line: u32,
) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    chunks[current].emit_string_const(key, line);
    chunks[current].emit_i32_const(value, line);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
}

fn normalize_slot_to_string(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    convert::emit_to_string(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, slot, line);
}

fn build_xattribute_object(
    chunks: &mut [Chunk],
    current: usize,
    name_slot: u16,
    value_slot: u16,
    obj_slot: u16,
    line: u32,
) {
    call_import(chunks, current, "ecma:object", "new", 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, obj_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    chunks[current].emit_string_const(KIND_KEY, line);
    chunks[current].emit_string_const(ATTR_KIND, line);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    set_field(chunks, current, obj_slot, NAME_KEY, name_slot, line);
    set_field(chunks, current, obj_slot, VALUE_KEY, value_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, obj_slot, line);
}

fn get_field(chunks: &mut [Chunk], current: usize, key: &str, line: u32) {
    chunks[current].emit_string_const(key, line);
    collections::emit_get(chunks, current, line);
}

fn add_xname_dotnet_aliases(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = chunks[current].alloc_scratch(3);
    let name_slot = base;
    let value_slot = base + 1;
    chunks[current].emit_op_u16(Op::LOCAL_SET, name_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, name_slot, line);
    vybe_compiler::primitives::xml::emit_local(chunks, current, 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, name_slot, line);
    chunks[current].emit_string_const("LocalName", line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, name_slot, line);
    chunks[current].emit_string_const("localname", line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, name_slot, line);
    vybe_compiler::primitives::xml::emit_namespace(chunks, current, 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, name_slot, line);
    chunks[current].emit_string_const("NamespaceName", line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, name_slot, line);
    chunks[current].emit_string_const("namespacename", line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, name_slot, line);
}

pub fn emit_xattribute_new(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let base = chunks[current].alloc_scratch(3);
    let name_slot = base;
    let value_slot = base + 1;
    let obj_slot = base + 2;

    chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, name_slot, line);
    normalize_slot_to_string(chunks, current, value_slot, line);
    build_xattribute_object(chunks, current, name_slot, value_slot, obj_slot, line);
}

pub fn emit_xcomment_new(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let base = chunks[current].alloc_scratch(2);
    let value_slot = base;
    let node_slot = base + 1;
    chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);
    normalize_slot_to_string(chunks, current, value_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    call_import(chunks, current, "web:dom-parser", "createComment", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, node_slot, line);
    set_field_const_str(chunks, current, node_slot, "__type", "XComment", line);
    set_field_const_str(chunks, current, node_slot, "NodeType", "Comment", line);
    set_field_const_str(chunks, current, node_slot, "nodetype", "Comment", line);
    set_field(chunks, current, node_slot, "Value", value_slot, line);
    set_field(chunks, current, node_slot, "value", value_slot, line);
    set_field(chunks, current, node_slot, "Data", value_slot, line);
    set_field(chunks, current, node_slot, "data", value_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, node_slot, line);
}

pub fn emit_xcdata_new(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let base = chunks[current].alloc_scratch(2);
    let value_slot = base;
    let node_slot = base + 1;
    chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);
    normalize_slot_to_string(chunks, current, value_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    call_import(
        chunks,
        current,
        "web:dom-parser",
        "createCDATASection",
        1,
        line,
    );
    chunks[current].emit_op_u16(Op::LOCAL_SET, node_slot, line);
    set_field_const_str(chunks, current, node_slot, "__type", "XCData", line);
    set_field_const_str(chunks, current, node_slot, "NodeType", "CDATA", line);
    set_field_const_str(chunks, current, node_slot, "nodetype", "CDATA", line);
    set_field(chunks, current, node_slot, "Value", value_slot, line);
    set_field(chunks, current, node_slot, "value", value_slot, line);
    set_field(chunks, current, node_slot, "Data", value_slot, line);
    set_field(chunks, current, node_slot, "data", value_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, node_slot, line);
}

pub fn emit_xprocessing_instruction_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = chunks[current].alloc_scratch(3);
    let target_slot = base;
    let data_slot = base + 1;
    let node_slot = base + 2;
    if argc >= 2 {
        chunks[current].emit_op_u16(Op::LOCAL_SET, data_slot, line);
    } else {
        chunks[current].emit_string_const("", line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, data_slot, line);
    }
    if argc >= 1 {
        chunks[current].emit_op_u16(Op::LOCAL_SET, target_slot, line);
    } else {
        chunks[current].emit_string_const("", line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, target_slot, line);
    }
    for _ in 2..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    normalize_slot_to_string(chunks, current, target_slot, line);
    normalize_slot_to_string(chunks, current, data_slot, line);
    call_import(chunks, current, "ecma:object", "new", 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, node_slot, line);
    set_field_const_str(
        chunks,
        current,
        node_slot,
        "__type",
        "XProcessingInstruction",
        line,
    );
    set_field_const_str(
        chunks,
        current,
        node_slot,
        "NodeType",
        "ProcessingInstruction",
        line,
    );
    set_field_const_str(
        chunks,
        current,
        node_slot,
        "nodetype",
        "ProcessingInstruction",
        line,
    );
    set_field_i32(chunks, current, node_slot, "nodeType", 7, line);
    set_field(chunks, current, node_slot, "nodeName", target_slot, line);
    set_field(chunks, current, node_slot, "textContent", data_slot, line);
    set_field(chunks, current, node_slot, "nodeValue", data_slot, line);
    set_field(chunks, current, node_slot, "Target", target_slot, line);
    set_field(chunks, current, node_slot, "target", target_slot, line);
    set_field(chunks, current, node_slot, "Data", data_slot, line);
    set_field(chunks, current, node_slot, "data", data_slot, line);
    set_field(chunks, current, node_slot, "Value", data_slot, line);
    set_field(chunks, current, node_slot, "value", data_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, node_slot, line);
}

pub fn emit_xelement_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let argc_u16 = argc as u16;
    let base = chunks[current].alloc_scratch(argc_u16 + 8);
    let elem_slot = base + argc_u16;
    let content_slot = elem_slot + 1;
    let type_slot = elem_slot + 2;
    let kind_slot = elem_slot + 3;
    let attr_name_slot = elem_slot + 4;
    let attr_value_slot = elem_slot + 5;
    let idx_slot = elem_slot + 6;
    let item_slot = elem_slot + 7;

    for index in (0..argc_u16).rev() {
        chunks[current].emit_op_u16(Op::LOCAL_SET, base + index, line);
    }

    if argc >= 1 {
        chunks[current].emit_op_u16(Op::LOCAL_GET, base, line);
    } else {
        chunks[current].emit_string_const("", line);
    }
    call_import(chunks, current, "web:dom-parser", "createElement", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, elem_slot, line);

    if argc < 2 {
        collections::emit_array_new(chunks, current, 0, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, content_slot, line);
        set_xml_children_fields(chunks, current, elem_slot, content_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
        return;
    }

    if argc == 2 {
        chunks[current].emit_op_u16(Op::LOCAL_GET, base + 1, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, content_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, content_slot, line);
        chunks[current].emit_op(Op::REF_IS_NULL, line);
        chunks[current].emit_if(line);
        chunks[current].emit_else(line);

        chunks[current].emit_op_u16(Op::LOCAL_GET, content_slot, line);
        call_import(chunks, current, "ecma:array", "isArray", 1, line);
        chunks[current].emit_if(line);

        let state = loops::emit_for_in_start(chunks, current, content_slot, idx_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, item_slot, line);
        append_xelement_content(
            chunks,
            current,
            elem_slot,
            item_slot,
            type_slot,
            kind_slot,
            attr_name_slot,
            attr_value_slot,
            line,
        );
        loops::emit_for_in_end(chunks, current, idx_slot, state, line);

        chunks[current].emit_else(line);
        append_xelement_content(
            chunks,
            current,
            elem_slot,
            content_slot,
            type_slot,
            kind_slot,
            attr_name_slot,
            attr_value_slot,
            line,
        );
        chunks[current].emit_end(line);
        chunks[current].emit_end(line);
    } else {
        for index in 1..argc_u16 {
            chunks[current].emit_op_u16(Op::LOCAL_GET, base + index, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, content_slot, line);
            append_xelement_content(
                chunks,
                current,
                elem_slot,
                content_slot,
                type_slot,
                kind_slot,
                attr_name_slot,
                attr_value_slot,
                line,
            );
        }
    }
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
}

fn append_xelement_content(
    chunks: &mut [Chunk],
    current: usize,
    elem_slot: u16,
    content_slot: u16,
    type_slot: u16,
    kind_slot: u16,
    attr_name_slot: u16,
    attr_value_slot: u16,
    line: u32,
) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, content_slot, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    chunks[current].emit_else(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, content_slot, line);
    call_import(chunks, current, "ecma:value", "typeof", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, type_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, type_slot, line);
    chunks[current].emit_string_const("object", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, content_slot, line);
    get_field(chunks, current, KIND_KEY, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, kind_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, kind_slot, line);
    chunks[current].emit_string_const(ATTR_KIND, line);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, content_slot, line);
    get_field(chunks, current, NAME_KEY, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, attr_name_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, content_slot, line);
    get_field(chunks, current, VALUE_KEY, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, attr_value_slot, line);
    normalize_slot_to_string(chunks, current, attr_value_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, attr_name_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, attr_value_slot, line);
    call_import(chunks, current, "web:dom-parser", "setAttribute", 3, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, content_slot, line);
    call_import(chunks, current, "web:dom-parser", "appendChild", 2, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);

    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, content_slot, line);
    call_import(chunks, current, "web:dom-parser", "createTextNode", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, kind_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, kind_slot, line);
    call_import(chunks, current, "web:dom-parser", "appendChild", 2, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

pub fn emit_xdocument_parse(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = chunks[current].alloc_scratch(2);
    let doc_slot = base;
    let root_slot = base + 1;
    call_import(chunks, current, "web:dom-parser", "parse", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, doc_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, doc_slot, line);
    emit_xdocument_root(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, root_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, root_slot, line);
    get_field(chunks, current, "nodeName", line);
    chunks[current].emit_string_const("parsererror", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    crate::emitter::core::exceptions::emit_throw_typed(
        chunks,
        current,
        "XmlException",
        "Data at the root level is invalid.",
        line,
    );
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, doc_slot, line);
    chunks[current].emit_end(line);
}

pub fn emit_xelement_parse(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_xdocument_parse(chunks, current, line);
    emit_xdocument_root(chunks, current, line);
}

pub fn emit_xdocument_load(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = chunks[current].alloc_scratch(3);
    let stream_slot = base;
    let buf_slot = base + 1;
    let len_slot = base + 2;
    chunks[current].emit_op_u16(Op::LOCAL_SET, stream_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, stream_slot, line);
    get_field(chunks, current, MS_BUF, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, buf_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, stream_slot, line);
    get_field(chunks, current, MS_LEN, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len_slot, line);

    call_import(chunks, current, "web:encoding", "decoderNew", 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, buf_slot, line);
    chunks[current].emit_f64_const(0.0, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_slot, line);
    call_import(chunks, current, "ecma:array", "slice", 3, line);
    call_import(
        chunks,
        current,
        "ecma:uint8array",
        "newFromIterable",
        1,
        line,
    );
    call_import(chunks, current, "web:encoding", "decode", 2, line);
    emit_xdocument_parse(chunks, current, line);
}

pub fn emit_xdocument_save(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = chunks[current].alloc_scratch(4);
    let doc_slot = base;
    let stream_slot = base + 1;
    let bytes_slot = base + 2;
    let len_slot = base + 3;
    chunks[current].emit_op_u16(Op::LOCAL_SET, stream_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, doc_slot, line);

    call_import(chunks, current, "web:encoding", "encoderNew", 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, doc_slot, line);
    emit_xml_to_string(chunks, current, line);
    call_import(chunks, current, "web:encoding", "encode", 2, line);
    call_import(chunks, current, "ecma:array", "from", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, bytes_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, bytes_slot, line);
    call_import(chunks, current, "ecma:array", "length", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len_slot, line);

    set_field(chunks, current, stream_slot, MS_BUF, bytes_slot, line);
    set_field(chunks, current, stream_slot, MS_LEN, len_slot, line);
    set_field(chunks, current, stream_slot, MS_POS, len_slot, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

pub fn emit_xdocument_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = chunks[current].alloc_scratch(argc as u16 + 3);
    let doc_slot = base + argc as u16;
    let item_slot = doc_slot + 1;
    let node_type_slot = doc_slot + 2;
    for index in (0..argc as u16).rev() {
        chunks[current].emit_op_u16(Op::LOCAL_SET, base + index, line);
    }
    call_import(chunks, current, "web:dom-parser", "createDocument", 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, doc_slot, line);
    if argc == 0 {
        chunks[current].emit_op_u16(Op::LOCAL_GET, doc_slot, line);
        return;
    }
    for index in 0..argc as u16 {
        chunks[current].emit_op_u16(Op::LOCAL_GET, base + index, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, item_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, item_slot, line);
        get_field(chunks, current, "nodeType", line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, node_type_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, node_type_slot, line);
        chunks[current].emit_i32_const(1, line);
        vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
        vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
        chunks[current].emit_if(line);
        set_field(
            chunks,
            current,
            doc_slot,
            "documentElement",
            item_slot,
            line,
        );
        chunks[current].emit_end(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, doc_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, item_slot, line);
        call_import(chunks, current, "web:dom-parser", "appendChild", 2, line);
        chunks[current].emit_op(Op::DROP, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_GET, doc_slot, line);
}

pub fn emit_xdocument_root(chunks: &mut [Chunk], current: usize, line: u32) {
    get_field(chunks, current, "documentElement", line);
}

pub fn emit_xml_first_node(chunks: &mut [Chunk], current: usize, line: u32) {
    get_field(chunks, current, "firstChild", line);
}

pub fn emit_xml_last_node(chunks: &mut [Chunk], current: usize, line: u32) {
    get_field(chunks, current, "lastChild", line);
}

pub fn emit_xml_parent(chunks: &mut [Chunk], current: usize, line: u32) {
    get_field(chunks, current, "parentNode", line);
}

pub fn emit_xml_ancestors(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = chunks[current].alloc_scratch(4);
    let elem_slot = base;
    let out_slot = base + 1;
    let current_slot = base + 2;
    let node_type_slot = base + 3;
    chunks[current].emit_op_u16(Op::LOCAL_SET, elem_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    emit_xml_parent(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, current_slot, line);

    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out_slot, line);

    for _ in 0..16 {
        chunks[current].emit_op_u16(Op::LOCAL_GET, current_slot, line);
        chunks[current].emit_op(Op::REF_IS_NULL, line);
        chunks[current].emit_if(line);
        chunks[current].emit_else(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, current_slot, line);
        get_field(chunks, current, "nodeType", line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, node_type_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, node_type_slot, line);
        chunks[current].emit_i32_const(1, line);
        vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
        vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
        chunks[current].emit_if(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, out_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, current_slot, line);
        collections::emit_push(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, current_slot, line);
        emit_xml_parent(chunks, current, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, current_slot, line);
        chunks[current].emit_else(line);
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, current_slot, line);
        chunks[current].emit_end(line);
        chunks[current].emit_end(line);
    }

    chunks[current].emit_op_u16(Op::LOCAL_GET, out_slot, line);
}

pub fn emit_xelement_name(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = chunks[current].alloc_scratch(2);
    let node_slot = base;
    let name_slot = base + 1;
    chunks[current].emit_op_u16(Op::LOCAL_SET, node_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, node_slot, line);
    vybe_compiler::primitives::xml::emit_node_name(chunks, current, 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, name_slot, line);
    fill_xname_namespace_from_xmlns(chunks, current, node_slot, name_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, name_slot, line);
    add_xname_dotnet_aliases(chunks, current, line);
}

fn fill_xname_namespace_from_xmlns(
    chunks: &mut [Chunk],
    current: usize,
    node_slot: u16,
    name_slot: u16,
    line: u32,
) {
    let base = chunks[current].alloc_scratch(4);
    let current_slot = base;
    let prefix_slot = base + 1;
    let attr_name_slot = base + 2;
    let namespace_slot = base + 3;

    chunks[current].emit_op_u16(Op::LOCAL_GET, name_slot, line);
    chunks[current].emit_string_const("prefix", line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, prefix_slot, line);
    chunks[current].emit_string_const("xmlns:", line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, prefix_slot, line);
    vybe_compiler::primitives::strings::emit_concat(&mut chunks[current], 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, attr_name_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, node_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, current_slot, line);

    for _ in 0..16 {
        chunks[current].emit_op_u16(Op::LOCAL_GET, current_slot, line);
        chunks[current].emit_op(Op::REF_IS_NULL, line);
        chunks[current].emit_if(line);
        chunks[current].emit_else(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, current_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, attr_name_slot, line);
        call_import(chunks, current, "web:dom-parser", "getAttribute", 2, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, namespace_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, namespace_slot, line);
        chunks[current].emit_op(Op::REF_IS_NULL, line);
        chunks[current].emit_if(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, current_slot, line);
        emit_xml_parent(chunks, current, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, current_slot, line);
        chunks[current].emit_else(line);
        set_field(
            chunks,
            current,
            name_slot,
            "namespaceURI",
            namespace_slot,
            line,
        );
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, current_slot, line);
        chunks[current].emit_end(line);
        chunks[current].emit_end(line);
    }
}

pub fn emit_xml_node_name(chunks: &mut [Chunk], current: usize, line: u32) {
    get_field(chunks, current, "nodeName", line);
}

pub fn emit_xml_value(chunks: &mut [Chunk], current: usize, line: u32) {
    let value_slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    call_import(chunks, current, "wasm:js-string", "test", 1, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    get_field(chunks, current, "textContent", line);
    chunks[current].emit_end(line);
}

pub fn emit_xml_to_string(chunks: &mut [Chunk], current: usize, line: u32) {
    call_import(chunks, current, "web:dom-parser", "toString", 1, line);
}

pub fn emit_xml_element(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_xml_elements(chunks, current, line);
    chunks[current].emit_f64_const(0.0, line);
    collections::emit_get(chunks, current, line);
}

pub fn emit_xml_child_elements(chunks: &mut [Chunk], current: usize, line: u32) {
    get_field(chunks, current, "children", line);
}

pub fn emit_xml_elements(chunks: &mut [Chunk], current: usize, line: u32) {
    call_import(
        chunks,
        current,
        "web:dom-parser",
        "getElementsByTagName",
        2,
        line,
    );
}

pub fn emit_xml_attribute(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = chunks[current].alloc_scratch(4);
    let name_slot = base;
    let elem_slot = base + 1;
    let value_slot = base + 2;
    let obj_slot = base + 3;
    chunks[current].emit_op_u16(Op::LOCAL_SET, name_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, elem_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, name_slot, line);
    call_import(chunks, current, "web:dom-parser", "getAttribute", 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, name_slot, line);
    vybe_compiler::primitives::strings::emit_to_lower(&mut chunks[current], line);
    call_import(chunks, current, "web:dom-parser", "getAttribute", 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunks[current].emit_else(line);
    build_xattribute_object(chunks, current, name_slot, value_slot, obj_slot, line);
    chunks[current].emit_end(line);
}

fn emit_xml_adapt_node_from_local(chunks: &mut [Chunk], current: usize, node_slot: u16, line: u32) {
    let base = chunks[current].alloc_scratch(4);
    let children_slot = base;
    let attrs_slot = base + 1;
    let child_count_slot = base + 2;
    let attr_count_slot = base + 3;
    chunks[current].emit_op_u16(Op::LOCAL_GET, node_slot, line);
    get_field(chunks, current, "children", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, children_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, children_slot, line);
    call_import(chunks, current, "ecma:array", "length", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, child_count_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, node_slot, line);
    get_field(chunks, current, "attributes", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, attrs_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, attrs_slot, line);
    call_import(chunks, current, "ecma:object", "keys", 1, line);
    call_import(chunks, current, "ecma:array", "length", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, attr_count_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, child_count_slot, line);
    chunks[current].emit_i32_const(0, line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, attr_count_slot, line);
    chunks[current].emit_i32_const(0, line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_AND, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, node_slot, line);
    get_field(chunks, current, "textContent", line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, node_slot, line);
    chunks[current].emit_end(line);
}

pub fn emit_xml_ps_get_member(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = chunks[current].alloc_scratch(9);
    let name_slot = base;
    let obj_slot = base + 1;
    let attr_slot = base + 2;
    let children_slot = base + 3;
    let matches_slot = base + 4;
    let idx_slot = base + 5;
    let item_slot = base + 6;
    let node_name_slot = base + 7;
    let len_slot = base + 8;

    chunks[current].emit_op_u16(Op::LOCAL_SET, name_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, obj_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    get_field(chunks, current, "documentElement", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, item_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, item_slot, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, item_slot, line);
    call_import(chunks, current, "wasm:js-undefined", "test", 1, line);
    chunks[current].emit_op(Op::I32_OR, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, item_slot, line);
    get_field(chunks, current, "nodeName", line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, name_slot, line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    emit_xml_adapt_node_from_local(chunks, current, item_slot, line);
    chunks[current].emit_else(line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, name_slot, line);
    call_import(chunks, current, "web:dom-parser", "getAttribute", 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, attr_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    get_field(chunks, current, "children", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, children_slot, line);
    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, matches_slot, line);

    let state = loops::emit_for_in_start(chunks, current, children_slot, idx_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, item_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, item_slot, line);
    get_field(chunks, current, "nodeName", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, node_name_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, node_name_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, name_slot, line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, matches_slot, line);
    emit_xml_adapt_node_from_local(chunks, current, item_slot, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);
    loops::emit_for_in_end(chunks, current, idx_slot, state, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, matches_slot, line);
    call_import(chunks, current, "ecma:array", "length", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, attr_slot, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_slot, line);
    chunks[current].emit_i32_const(0, line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, attr_slot, line);
    chunks[current].emit_else(line);
    chunks[current].emit_array_new_fixed(0, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, children_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, children_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, attr_slot, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    let state = loops::emit_for_in_start(chunks, current, matches_slot, idx_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, item_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, children_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, item_slot, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    loops::emit_for_in_end(chunks, current, idx_slot, state, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, children_slot, line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_slot, line);
    chunks[current].emit_i32_const(0, line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_slot, line);
    chunks[current].emit_i32_const(1, line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, matches_slot, line);
    chunks[current].emit_f64_const(0.0, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, matches_slot, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

pub fn emit_xml_ps_set_member(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = chunks[current].alloc_scratch(11);
    let value_slot = base;
    let name_slot = base + 1;
    let obj_slot = base + 2;
    let attr_slot = base + 3;
    let children_slot = base + 4;
    let idx_slot = base + 5;
    let item_slot = base + 6;
    let child_nodes_slot = base + 7;
    let first_child_slot = base + 8;
    let len_slot = base + 9;
    let text_node_slot = base + 10;
    chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, name_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, obj_slot, line);
    normalize_slot_to_string(chunks, current, value_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, name_slot, line);
    chunks[current].emit_string_const("InnerText", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    emit_xml_replace_text(
        chunks,
        current,
        obj_slot,
        value_slot,
        child_nodes_slot,
        first_child_slot,
        len_slot,
        text_node_slot,
        line,
    );
    chunks[current].emit_else(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, name_slot, line);
    call_import(chunks, current, "web:dom-parser", "getAttribute", 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, attr_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, attr_slot, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, name_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    call_import(chunks, current, "web:dom-parser", "setAttribute", 3, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    get_field(chunks, current, "children", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, children_slot, line);
    let state = loops::emit_for_in_start(chunks, current, children_slot, idx_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, item_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, item_slot, line);
    get_field(chunks, current, "nodeName", line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, name_slot, line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    emit_xml_replace_text(
        chunks,
        current,
        item_slot,
        value_slot,
        child_nodes_slot,
        first_child_slot,
        len_slot,
        text_node_slot,
        line,
    );
    chunks[current].emit_end(line);
    loops::emit_for_in_end(chunks, current, idx_slot, state, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

fn emit_xml_replace_text(
    chunks: &mut [Chunk],
    current: usize,
    node_slot: u16,
    value_slot: u16,
    child_nodes_slot: u16,
    first_child_slot: u16,
    len_slot: u16,
    text_node_slot: u16,
    line: u32,
) {
    set_field(chunks, current, node_slot, "textContent", value_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, node_slot, line);
    get_field(chunks, current, "childNodes", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, child_nodes_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, child_nodes_slot, line);
    call_import(chunks, current, "ecma:array", "length", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_slot, line);
    chunks[current].emit_i32_const(0, line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, child_nodes_slot, line);
    chunks[current].emit_f64_const(0.0, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, first_child_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, node_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, first_child_slot, line);
    call_import(chunks, current, "web:dom-parser", "removeChild", 2, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    call_import(chunks, current, "web:dom-parser", "createTextNode", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, text_node_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, node_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, text_node_slot, line);
    call_import(chunks, current, "web:dom-parser", "appendChild", 2, line);
    chunks[current].emit_op(Op::DROP, line);
}

pub fn emit_xml_inner_xml(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = chunks[current].alloc_scratch(5);
    let obj_slot = base;
    let children_slot = base + 1;
    let out_slot = base + 2;
    let idx_slot = base + 3;
    let item_slot = base + 4;
    chunks[current].emit_op_u16(Op::LOCAL_SET, obj_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    get_field(chunks, current, "children", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, children_slot, line);
    chunks[current].emit_string_const("", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out_slot, line);
    let state = loops::emit_for_in_start(chunks, current, children_slot, idx_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, item_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, item_slot, line);
    emit_xml_to_string(chunks, current, line);
    vybe_compiler::primitives::strings::emit_str_concat(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out_slot, line);
    loops::emit_for_in_end(chunks, current, idx_slot, state, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out_slot, line);
}

pub fn emit_xml_has_child_nodes(chunks: &mut [Chunk], current: usize, line: u32) {
    let value_slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    call_import(chunks, current, "ecma:value", "typeof", 1, line);
    chunks[current].emit_string_const("object", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    get_field(chunks, current, "children", line);
    call_import(chunks, current, "ecma:array", "length", 1, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_NE, line);
    chunks[current].emit_else(line);
    chunks[current].emit_bool_const(false, line);
    chunks[current].emit_end(line);
}

pub fn emit_xml_attributes(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = chunks[current].alloc_scratch(8);
    let obj_slot = base;
    let attrs_slot = base + 1;
    let keys_slot = base + 2;
    let out_slot = base + 3;
    let idx_slot = base + 4;
    let name_slot = base + 5;
    let value_slot = base + 6;
    let attr_obj_slot = base + 7;
    chunks[current].emit_op_u16(Op::LOCAL_SET, obj_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    get_field(chunks, current, "attributes", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, attrs_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, attrs_slot, line);
    call_import(chunks, current, "ecma:object", "keys", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, keys_slot, line);
    call_import(chunks, current, "ecma:object", "new", 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out_slot, line);
    let state = loops::emit_for_in_start(chunks, current, keys_slot, idx_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, name_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, attrs_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, name_slot, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);
    build_xattribute_object(chunks, current, name_slot, value_slot, attr_obj_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, attr_obj_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, name_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, attr_obj_slot, line);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    loops::emit_for_in_end(chunks, current, idx_slot, state, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out_slot, line);
}

fn emit_xpath_selector(
    chunks: &mut [Chunk],
    current: usize,
    source_slot: u16,
    dest_slot: u16,
    line: u32,
) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, source_slot, line);
    chunks[current].emit_string_const("//", line);
    chunks[current].emit_string_const("", line);
    call_import(chunks, current, "ecma:string", "replaceAll", 3, line);
    chunks[current].emit_string_const("[@", line);
    chunks[current].emit_string_const("[", line);
    call_import(chunks, current, "ecma:string", "replaceAll", 3, line);
    chunks[current].emit_string_const("[last()]", line);
    chunks[current].emit_string_const("", line);
    call_import(chunks, current, "ecma:string", "replaceAll", 3, line);
    chunks[current].emit_string_const("[1]", line);
    chunks[current].emit_string_const("", line);
    call_import(chunks, current, "ecma:string", "replaceAll", 3, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, dest_slot, line);
}

pub fn emit_xml_select_nodes(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = chunks[current].alloc_scratch(3);
    let xpath_slot = base;
    let obj_slot = base + 1;
    let selector_slot = base + 2;
    chunks[current].emit_op_u16(Op::LOCAL_SET, xpath_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, obj_slot, line);
    emit_xpath_selector(chunks, current, xpath_slot, selector_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, selector_slot, line);
    call_import(
        chunks,
        current,
        "web:dom-parser",
        "querySelectorAll",
        2,
        line,
    );
}

pub fn emit_xml_select_single_node(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = chunks[current].alloc_scratch(5);
    let xpath_slot = base;
    let obj_slot = base + 1;
    let selector_slot = base + 2;
    let nodes_slot = base + 3;
    let len_slot = base + 4;
    chunks[current].emit_op_u16(Op::LOCAL_SET, xpath_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, obj_slot, line);
    emit_xpath_selector(chunks, current, xpath_slot, selector_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, xpath_slot, line);
    chunks[current].emit_string_const("[last()]", line);
    call_import(chunks, current, "ecma:string", "indexOf", 2, line);
    chunks[current].emit_i32_const(-1, line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, selector_slot, line);
    call_import(
        chunks,
        current,
        "web:dom-parser",
        "querySelectorAll",
        2,
        line,
    );
    chunks[current].emit_op_u16(Op::LOCAL_SET, nodes_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, nodes_slot, line);
    call_import(chunks, current, "ecma:array", "length", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_slot, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_SUB, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, nodes_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_slot, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, selector_slot, line);
    call_import(chunks, current, "web:dom-parser", "querySelector", 2, line);
    chunks[current].emit_end(line);
}

pub fn emit_xml_load_xml(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = chunks[current].alloc_scratch(4);
    let xml_slot = base;
    let doc_slot = base + 1;
    let parsed_slot = base + 2;
    let field_slot = base + 3;
    chunks[current].emit_op_u16(Op::LOCAL_SET, xml_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, doc_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, xml_slot, line);
    emit_xdocument_parse(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, parsed_slot, line);
    for field in ["documentElement", "childNodes", "children", "textContent"] {
        chunks[current].emit_op_u16(Op::LOCAL_GET, parsed_slot, line);
        get_field(chunks, current, field, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, field_slot, line);
        set_field(chunks, current, doc_slot, field, field_slot, line);
    }
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

pub fn emit_xml_create_element(chunks: &mut [Chunk], current: usize, line: u32) {
    call_import(chunks, current, "web:dom-parser", "createElement", 2, line);
}

pub fn emit_xml_set_attribute(chunks: &mut [Chunk], current: usize, line: u32) {
    call_import(chunks, current, "web:dom-parser", "setAttribute", 3, line);
}

pub fn emit_xml_append_child(chunks: &mut [Chunk], current: usize, line: u32) {
    call_import(chunks, current, "web:dom-parser", "appendChild", 2, line);
}

pub fn emit_xml_remove_child(chunks: &mut [Chunk], current: usize, line: u32) {
    call_import(chunks, current, "web:dom-parser", "removeChild", 2, line);
}

pub fn emit_xml_clone_node(chunks: &mut [Chunk], current: usize, line: u32) {
    call_import(chunks, current, "web:dom-parser", "cloneNode", 2, line);
}

pub fn emit_attribute_value(chunks: &mut [Chunk], current: usize, line: u32) {
    get_field(chunks, current, VALUE_KEY, line);
}

pub fn emit_xelement_add(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = chunks[current].alloc_scratch(9);
    let elem_slot = base;
    let content_slot = base + 1;
    let type_slot = base + 2;
    let kind_slot = base + 3;
    let attr_name_slot = base + 4;
    let attr_value_slot = base + 5;
    let idx_slot = base + 6;
    let item_slot = base + 7;
    chunks[current].emit_op_u16(Op::LOCAL_SET, content_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, elem_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, content_slot, line);
    call_import(chunks, current, "ecma:array", "isArray", 1, line);
    chunks[current].emit_if(line);
    let state = loops::emit_for_in_start(chunks, current, content_slot, idx_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, item_slot, line);
    append_xelement_content(
        chunks,
        current,
        elem_slot,
        item_slot,
        type_slot,
        kind_slot,
        attr_name_slot,
        attr_value_slot,
        line,
    );
    loops::emit_for_in_end(chunks, current, idx_slot, state, line);
    chunks[current].emit_else(line);
    append_xelement_content(
        chunks,
        current,
        elem_slot,
        content_slot,
        type_slot,
        kind_slot,
        attr_name_slot,
        attr_value_slot,
        line,
    );
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
}

pub fn emit_xelement_remove(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = chunks[current].alloc_scratch(2);
    let elem_slot = base;
    let parent_slot = base + 1;
    chunks[current].emit_op_u16(Op::LOCAL_SET, elem_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    get_field(chunks, current, "parentNode", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, parent_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, parent_slot, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, parent_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    call_import(chunks, current, "web:dom-parser", "removeChild", 2, line);
    chunks[current].emit_end(line);
}

pub fn emit_xelement_replace_with(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = chunks[current].alloc_scratch(3);
    let old_slot = base;
    let new_slot = base + 1;
    let parent_slot = base + 2;
    chunks[current].emit_op_u16(Op::LOCAL_SET, new_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, old_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, old_slot, line);
    get_field(chunks, current, "parentNode", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, parent_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, parent_slot, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, parent_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, new_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, old_slot, line);
    call_import(chunks, current, "web:dom-parser", "replaceChild", 3, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, new_slot, line);
}

pub fn emit_xelement_replace_nodes(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = chunks[current].alloc_scratch(9);
    let elem_slot = base;
    let content_slot = base + 1;
    let type_slot = base + 2;
    let kind_slot = base + 3;
    let attr_name_slot = base + 4;
    let attr_value_slot = base + 5;
    let idx_slot = base + 6;
    let item_slot = base + 7;
    let children_slot = base + 8;
    chunks[current].emit_op_u16(Op::LOCAL_SET, content_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, elem_slot, line);
    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, children_slot, line);
    set_xml_children_fields(chunks, current, elem_slot, children_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, content_slot, line);
    call_import(chunks, current, "ecma:array", "isArray", 1, line);
    chunks[current].emit_if(line);
    let state = loops::emit_for_in_start(chunks, current, content_slot, idx_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, item_slot, line);
    append_xelement_content(
        chunks,
        current,
        elem_slot,
        item_slot,
        type_slot,
        kind_slot,
        attr_name_slot,
        attr_value_slot,
        line,
    );
    loops::emit_for_in_end(chunks, current, idx_slot, state, line);
    chunks[current].emit_else(line);
    append_xelement_content(
        chunks,
        current,
        elem_slot,
        content_slot,
        type_slot,
        kind_slot,
        attr_name_slot,
        attr_value_slot,
        line,
    );
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
}

pub fn emit_xelement_set_element_value(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = chunks[current].alloc_scratch(11);
    let elem_slot = base;
    let name_slot = base + 1;
    let value_slot = base + 2;
    let child_slot = base + 3;
    let type_slot = base + 4;
    let kind_slot = base + 5;
    let attr_name_slot = base + 6;
    let attr_value_slot = base + 7;
    let idx_slot = base + 8;
    let item_slot = base + 9;
    let children_slot = base + 10;

    chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, name_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, elem_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, name_slot, line);
    emit_xml_element(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, child_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, child_slot, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, name_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    emit_xelement_new(chunks, current, 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, child_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, child_slot, line);
    call_import(chunks, current, "web:dom-parser", "appendChild", 2, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_else(line);
    normalize_slot_to_string(chunks, current, value_slot, line);
    emit_xml_replace_text(
        chunks,
        current,
        child_slot,
        value_slot,
        children_slot,
        type_slot,
        kind_slot,
        attr_name_slot,
        line,
    );
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
}

pub fn emit_xelement_set_attribute_value(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = chunks[current].alloc_scratch(3);
    let elem_slot = base;
    let name_slot = base + 1;
    let value_slot = base + 2;
    chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, name_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, elem_slot, line);
    normalize_slot_to_string(chunks, current, value_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, name_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    call_import(chunks, current, "web:dom-parser", "setAttribute", 3, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
}

pub fn emit_xnode_deep_equals(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = chunks[current].alloc_scratch(4);
    let right_slot = base;
    let left_slot = base + 1;
    let right_text_slot = base + 2;
    let left_text_slot = base + 3;
    chunks[current].emit_op_u16(Op::LOCAL_SET, right_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, left_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, left_slot, line);
    emit_xml_to_string(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, left_text_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, right_slot, line);
    emit_xml_to_string(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, right_text_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, left_text_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, right_text_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
}
