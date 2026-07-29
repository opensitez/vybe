use vybe_bytecode::opcode::Op;
use vybe_bytecode::Chunk;
use vybe_compiler::compiler::{collections, loops};

const ATTR_KIND: &str = "XAttribute";
const KIND_KEY: &str = "__dotnet_xml_kind";
const NAME_KEY: &str = "name";
const VALUE_KEY: &str = "value";

fn call_import(
    chunks: &mut [Chunk],
    current: usize,
    module: &str,
    name: &str,
    argc: u8,
    line: u32,
) {
    let idx = chunks[current].add_import(module, name);
    chunks[current].emit_op_u16(Op::CALL_IMPORT, idx, line);
    chunks[current].emit(argc, line);
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
    vybe_compiler::compiler::xml::emit_local(chunks, current, 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, name_slot, line);
    chunks[current].emit_string_const("LocalName", line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, name_slot, line);
    vybe_compiler::compiler::xml::emit_namespace(chunks, current, 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, name_slot, line);
    chunks[current].emit_string_const("NamespaceName", line);
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
    build_xattribute_object(chunks, current, name_slot, value_slot, obj_slot, line);
}

pub fn emit_xelement_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = chunks[current].alloc_scratch(9);
    let name_slot = base;
    let content_slot = base + 1;
    let elem_slot = base + 2;
    let type_slot = base + 3;
    let kind_slot = base + 4;
    let attr_name_slot = base + 5;
    let attr_value_slot = base + 6;
    let idx_slot = base + 7;
    let item_slot = base + 8;

    if argc >= 2 {
        chunks[current].emit_op_u16(Op::LOCAL_SET, content_slot, line);
    }
    if argc >= 1 {
        chunks[current].emit_op_u16(Op::LOCAL_SET, name_slot, line);
    } else {
        chunks[current].emit_string_const("", line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, name_slot, line);
    }
    for _ in 2..argc {
        chunks[current].emit_op(Op::DROP, line);
    }

    chunks[current].emit_op_u16(Op::LOCAL_GET, name_slot, line);
    call_import(chunks, current, "web:dom-parser", "createElement", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, elem_slot, line);

    if argc < 2 {
        chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
        return;
    }

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
    vybe_compiler::compiler::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_compiler::compiler::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, content_slot, line);
    get_field(chunks, current, KIND_KEY, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, kind_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, kind_slot, line);
    chunks[current].emit_string_const(ATTR_KIND, line);
    vybe_compiler::compiler::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_compiler::compiler::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, content_slot, line);
    get_field(chunks, current, NAME_KEY, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, attr_name_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, content_slot, line);
    get_field(chunks, current, VALUE_KEY, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, attr_value_slot, line);
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
    call_import(chunks, current, "web:dom-parser", "parse", 1, line);
}

pub fn emit_xdocument_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc == 0 {
        call_import(chunks, current, "web:dom-parser", "createDocument", 0, line);
    } else {
        for _ in 1..argc {
            chunks[current].emit_op(Op::DROP, line);
        }
        emit_xdocument_parse(chunks, current, line);
    }
}

pub fn emit_xdocument_root(chunks: &mut [Chunk], current: usize, line: u32) {
    get_field(chunks, current, "documentElement", line);
}

pub fn emit_xelement_name(chunks: &mut [Chunk], current: usize, line: u32) {
    vybe_compiler::compiler::xml::emit_node_name(chunks, current, 1, line);
    add_xname_dotnet_aliases(chunks, current, line);
}

pub fn emit_xml_value(chunks: &mut [Chunk], current: usize, line: u32) {
    get_field(chunks, current, "textContent", line);
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
    build_xattribute_object(chunks, current, name_slot, value_slot, obj_slot, line);
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
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    chunks[current].emit_string_const("childNodes", line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, children_slot, line);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
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

pub fn emit_xelement_set_attribute_value(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = chunks[current].alloc_scratch(3);
    let elem_slot = base;
    let name_slot = base + 1;
    let value_slot = base + 2;
    chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, name_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, elem_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, name_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    call_import(chunks, current, "web:dom-parser", "setAttribute", 3, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
}
