//! Shared XML value helpers.
//!
//! XML names are the compatibility hinge between language surfaces:
//! Go exposes `xml.Name{Space, Local}`, .NET exposes `XName.NamespaceName`
//! / `LocalName`, Java exposes `QName.{namespaceURI, localPart, prefix}`,
//! and DOM-style APIs expose `namespaceURI`, `localName`, and `prefix`.
//!
//! This module gives those adapters one portable backing shape instead of
//! each frontend inventing its own "almost a QName" object.

use crate::primitives::collections;
use crate::primitives::ops;
use vybe_bytecode::Chunk;
use vybe_bytecode::namespaces::{self, NamespaceNode, Subtree};
use vybe_bytecode::opcode::Op;

pub const XML_NAME_TYPE: &str = "XmlName";
pub const FIELD_TYPE: &str = "__type";
pub const FIELD_LOCAL: &str = "localName";
pub const FIELD_NAMESPACE: &str = "namespaceURI";
pub const FIELD_PREFIX: &str = "prefix";
pub const FIELD_NODE_NAME: &str = "nodeName";

/// Register the portable XML helper namespace (`xml.name`, `xml.local`, ...).
///
/// Language frontends should target these canonical leaves directly. Source
/// syntax like Go `xml.Name`, .NET `XName`, Java `QName`, and VB XML literals
/// can still normalize however they need, but the runtime name object shape is
/// shared here.
pub fn register_namespace_tree() {
    let mut tree = Subtree::new();
    for (name, emit) in [
        ("name", "xml.name"),
        ("local", "xml.local"),
        ("namespace", "xml.namespace"),
        ("prefix", "xml.prefix"),
        ("qualified", "xml.qualified"),
        ("equal", "xml.equal"),
        ("from_dom_node", "xml.from_dom_node"),
        ("node_name", "xml.node_name"),
        ("parse", "xml.parse"),
        ("load", "xml.load"),
        ("save", "xml.save"),
        ("elements", "xml.elements"),
    ] {
        tree.insert(name.into(), NamespaceNode::CommonEmit(emit.into()));
    }
    namespaces::register_namespace_tree("xml", NamespaceNode::Namespace(tree));
}

fn emit_object_new(chunks: &mut [Chunk], current: usize, line: u32) {
    let idx = chunks[current].add_import("ecma:object", "new");
    chunks[current].emit_call(idx, 0, line);
}

fn emit_set_literal_field(
    chunks: &mut [Chunk],
    current: usize,
    object_slot: u16,
    key: &str,
    value: &str,
    line: u32,
) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, object_slot, line);
    chunks[current].emit_string_const(key, line);
    chunks[current].emit_string_const(value, line);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
}

fn emit_set_slot_field(
    chunks: &mut [Chunk],
    current: usize,
    object_slot: u16,
    key: &str,
    value_slot: u16,
    line: u32,
) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, object_slot, line);
    chunks[current].emit_string_const(key, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
}

fn emit_field_get(chunks: &mut [Chunk], current: usize, key: &str, line: u32) {
    chunks[current].emit_string_const(key, line);
    collections::emit_get(chunks, current, line);
}

fn emit_field_get_or_empty(chunks: &mut [Chunk], current: usize, key: &str, line: u32) {
    let value_slot = chunks[current].alloc_scratch(1);
    emit_field_get(chunks, current, key, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_string_const("", line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunks[current].emit_end(line);
}

fn emit_normalize_slot_to_empty(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if(line);
    chunk.emit_string_const("", line);
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
    chunk.emit_end(line);
}

/// Build a shared XML qualified-name object.
///
/// Contract:
/// - `localName` is the required semantic local part.
/// - missing namespace and prefix are stored as `""`, not `null`.
/// - equality compares `(namespaceURI, localName)`; prefix is lexical only.
///
/// Stack: `[namespace_uri, local_name, prefix] -> [XmlName]`.
pub fn emit_name(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let prefix_slot = chunks[current].alloc_scratch(4);
    let local_slot = prefix_slot + 1;
    let namespace_slot = prefix_slot + 2;
    let object_slot = prefix_slot + 3;

    chunks[current].emit_op_u16(Op::LOCAL_SET, prefix_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, local_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, namespace_slot, line);

    emit_normalize_slot_to_empty(&mut chunks[current], namespace_slot, line);
    emit_normalize_slot_to_empty(&mut chunks[current], prefix_slot, line);

    emit_object_new(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, object_slot, line);

    emit_set_literal_field(
        chunks,
        current,
        object_slot,
        FIELD_TYPE,
        XML_NAME_TYPE,
        line,
    );
    emit_set_slot_field(
        chunks,
        current,
        object_slot,
        FIELD_NAMESPACE,
        namespace_slot,
        line,
    );
    emit_set_slot_field(chunks, current, object_slot, FIELD_LOCAL, local_slot, line);
    emit_set_slot_field(
        chunks,
        current,
        object_slot,
        FIELD_PREFIX,
        prefix_slot,
        line,
    );

    chunks[current].emit_op_u16(Op::LOCAL_GET, object_slot, line);
}

/// Build an `XmlName` from a DOM Element-like node.
///
/// Stack: `[node] -> [XmlName]`.
pub fn emit_from_dom_node(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let node_slot = chunks[current].alloc_scratch(4);
    let namespace_slot = node_slot + 1;
    let local_slot = node_slot + 2;
    let prefix_slot = node_slot + 3;

    chunks[current].emit_op_u16(Op::LOCAL_SET, node_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, node_slot, line);
    emit_field_get_or_empty(chunks, current, FIELD_NAMESPACE, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, namespace_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, node_slot, line);
    emit_field_get_or_empty(chunks, current, FIELD_LOCAL, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, local_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, node_slot, line);
    emit_field_get_or_empty(chunks, current, FIELD_PREFIX, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, prefix_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, namespace_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, local_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, prefix_slot, line);
    emit_name(chunks, current, 3, line);
}

/// Alias for callers that read the XML name of a DOM node.
///
/// Stack: `[node] -> [XmlName]`.
pub fn emit_node_name(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_from_dom_node(chunks, current, argc, line);
}

/// Stack: `[XmlName] -> [localName]`.
pub fn emit_local(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    emit_field_get(chunks, current, FIELD_LOCAL, line);
}

/// Stack: `[XmlName] -> [namespaceURI]`.
pub fn emit_namespace(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    emit_field_get(chunks, current, FIELD_NAMESPACE, line);
}

/// Stack: `[XmlName] -> [prefix]`.
pub fn emit_prefix(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    emit_field_get(chunks, current, FIELD_PREFIX, line);
}

/// XML name equality: namespace URI + local part. Prefix is lexical sugar.
///
/// Stack: `[left, right] -> [bool]`.
pub fn emit_equal(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let right_slot = chunks[current].alloc_scratch(2);
    let left_slot = right_slot + 1;

    chunks[current].emit_op_u16(Op::LOCAL_SET, right_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, left_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, left_slot, line);
    emit_field_get_or_empty(chunks, current, FIELD_NAMESPACE, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, right_slot, line);
    emit_field_get_or_empty(chunks, current, FIELD_NAMESPACE, line);
    let compare = chunks[current].add_import("wasm:js-string", "compare");
    chunks[current].emit_call(compare, 2, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if_value(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, left_slot, line);
    emit_field_get_or_empty(chunks, current, FIELD_LOCAL, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, right_slot, line);
    emit_field_get_or_empty(chunks, current, FIELD_LOCAL, line);
    let compare = chunks[current].add_import("wasm:js-string", "compare");
    chunks[current].emit_call(compare, 2, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    ops::emit_i32_to_bool(&mut chunks[current], line);

    chunks[current].emit_else(line);
    chunks[current].emit_bool_const(false, line);
    chunks[current].emit_end(line);
}

/// Stack: `[XmlName] -> [qualified string]`.
///
/// This is intentionally the XML lexical form (`prefix:local` when prefix is
/// present) rather than `{namespace}local`; it matches how source XML names are
/// written and keeps display/debug output language-neutral.
pub fn emit_qualified(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let name_slot = chunks[current].alloc_scratch(2);
    let prefix_slot = name_slot + 1;

    chunks[current].emit_op_u16(Op::LOCAL_SET, name_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, name_slot, line);
    emit_field_get(chunks, current, FIELD_PREFIX, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, prefix_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, prefix_slot, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if_value(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, name_slot, line);
    emit_field_get(chunks, current, FIELD_LOCAL, line);

    chunks[current].emit_else(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, prefix_slot, line);
    chunks[current].emit_string_const("", line);
    let eq = chunks[current].add_import("wasm:js-string", "compare");
    chunks[current].emit_call(eq, 2, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if_value(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, name_slot, line);
    emit_field_get(chunks, current, FIELD_LOCAL, line);

    chunks[current].emit_else(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, prefix_slot, line);
    chunks[current].emit_string_const(":", line);
    let concat = chunks[current].add_import("wasm:js-string", "concat");
    chunks[current].emit_call(concat, 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, name_slot, line);
    emit_field_get(chunks, current, FIELD_LOCAL, line);
    let concat = chunks[current].add_import("wasm:js-string", "concat");
    chunks[current].emit_call(concat, 2, line);

    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

/// Stack: `[xml_text] -> [document/node]`.
pub fn emit_parse(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let idx = chunks[current].add_import("web:dom-parser", "parse");
    chunks[current].emit_call(idx, 1, line);
}

/// Stack: `[path_or_text] -> [document/node]`.
pub fn emit_load(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let idx = chunks[current].add_import("web:dom-parser", "load");
    chunks[current].emit_call(idx, 1, line);
}

/// Stack: `[document/node] -> [xml_text]`.
pub fn emit_save(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let idx = chunks[current].add_import("web:dom-parser", "toString");
    chunks[current].emit_call(idx, 1, line);
}

/// Stack: `[document/node, tag_name] -> [element_sequence]`.
pub fn emit_elements(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let idx = chunks[current].add_import("web:dom-parser", "getElementsByTagName");
    chunks[current].emit_call(idx, 2, line);
}

#[cfg(test)]
mod tests {
    use super::*;
    use vybe_bytecode::namespaces::{NamespaceNode, clear_registry_for_tests, registry_read};

    /// This asserts REGISTRATION, so it checks the registry directly rather
    /// than resolving through it — resolution is the compiler's concern and
    /// lives there.
    #[test]
    fn xml_helpers_register_common_namespace_tree() {
        clear_registry_for_tests();
        register_namespace_tree();

        let guard = registry_read();
        let Some(NamespaceNode::Namespace(children)) = guard.tree.get("xml") else {
            panic!("xml root not registered");
        };
        for member in ["name", "node_name", "elements"] {
            assert!(children.contains_key(member), "missing xml.{member}");
        }
    }
}
