//! `web:dom-parser` DOM host — `createDocument`
//! (DOM Living Standard §4.5.1 `DOMImplementation.createDocument`).

use std::sync::Arc;
use vybe_bytecode::{Chunk, Op, VM, Value};
use vybe_emitter::platforms::register_platforms_all;

fn call(module: &str, name: &str, args: Vec<Value>) -> Value {
    let mut chunk = Chunk::new("<test>");
    let import_idx = chunk.add_import(module, name);
    let argc = args.len() as u8;
    for v in args {
        let k = chunk.add_constant(v);
        chunk.emit_op_u16(Op::CONST, k, 0);
    }
    chunk.emit_op_u16(Op::CALL_IMPORT, import_idx, 0);
    chunk.emit(argc, 0);
    chunk.emit_op(Op::RETURN, 0);
    let mut vm = VM::new();
    register_platforms_all(&mut vm);
    vm.run(vec![chunk]).expect("VM run failed")
}

fn prop(v: &Value, key: &str) -> Value {
    match v {
        Value::Object(o) => o
            .lock()
            .unwrap()
            .properties
            .get(key)
            .cloned()
            .unwrap_or(Value::Null),
        _ => Value::Null,
    }
}

/// No `qualifiedName` → a fresh empty document (nodeType 9), no root element.
#[test]
fn create_document_empty() {
    let doc = call(
        "web:dom-parser",
        "createDocument",
        vec![Value::Null, Value::Null, Value::Null],
    );
    assert!(matches!(prop(&doc, "nodeType"), Value::I32(9)));
    assert!(matches!(prop(&doc, "__type"), Value::String(s) if &*s == "Document"));
    assert!(matches!(prop(&doc, "documentElement"), Value::Null));
}

/// A non-empty `qualifiedName` creates and appends a document element.
#[test]
fn create_document_with_root() {
    let doc = call(
        "web:dom-parser",
        "createDocument",
        vec![Value::Null, Value::String(Arc::from("root")), Value::Null],
    );
    let root = prop(&doc, "documentElement");
    assert!(matches!(prop(&root, "nodeType"), Value::I32(1)));
    assert!(matches!(prop(&root, "tagName"), Value::String(s) if &*s == "root"));
}

/// The namespace argument is recorded on the created root element.
#[test]
fn create_document_with_namespace() {
    let doc = call(
        "web:dom-parser",
        "createDocument",
        vec![
            Value::String(Arc::from("http://example.com/ns")),
            Value::String(Arc::from("ns:root")),
            Value::Null,
        ],
    );
    let root = prop(&doc, "documentElement");
    assert!(
        matches!(prop(&root, "namespaceURI"), Value::String(s) if &*s == "http://example.com/ns")
    );
}

/// `createCDATASection` (DOM §4.5 `Document.createCDATASection`) yields a
/// CDATA node (nodeType 4) carrying the raw text verbatim.
#[test]
fn create_cdata_section() {
    let doc = call(
        "web:dom-parser",
        "createDocument",
        vec![Value::Null, Value::Null, Value::Null],
    );
    let cdata = call(
        "web:dom-parser",
        "createCDATASection",
        vec![doc, Value::String(Arc::from("raw<data>"))],
    );
    assert!(matches!(prop(&cdata, "nodeType"), Value::I32(4)));
    assert!(matches!(prop(&cdata, "textContent"), Value::String(s) if &*s == "raw<data>"));
    assert!(matches!(prop(&cdata, "nodeName"), Value::String(s) if &*s == "#cdata-section"));
}

/// A parsed element's `childNodes` list exposes a live `length` (DOM §4.2.10
/// `NodeList.length`).
#[test]
fn node_list_length() {
    let doc = call(
        "web:dom-parser",
        "parse",
        vec![Value::String(Arc::from("<root><a/><b/></root>"))],
    );
    let root = prop(&doc, "documentElement");
    let children = prop(&root, "childNodes");
    assert!(matches!(prop(&children, "length"), Value::I32(2)));
}

/// `getElementById` (DOM §4.5) matches the element whose `id` attribute equals
/// the argument, even when the id came from parsing (not `setAttribute`).
#[test]
fn get_element_by_id_matches_attribute() {
    let doc = call(
        "web:dom-parser",
        "parse",
        vec![Value::String(Arc::from(
            r#"<root><item id="main">x</item></root>"#,
        ))],
    );
    let found = call(
        "web:dom-parser",
        "getElementById",
        vec![doc, Value::String(Arc::from("main"))],
    );
    assert!(matches!(prop(&found, "nodeType"), Value::I32(1)));
    assert!(matches!(prop(&found, "textContent"), Value::String(s) if &*s == "x"));
}

/// Appending a DocumentFragment moves its children into the parent and leaves
/// the fragment empty (DOM §4.2.1 pre-insert step for fragments).
#[test]
fn append_fragment_moves_children() {
    let doc = call(
        "web:dom-parser",
        "createDocument",
        vec![Value::Null, Value::Null, Value::Null],
    );
    let frag = call(
        "web:dom-parser",
        "createDocumentFragment",
        vec![doc.clone()],
    );
    call(
        "web:dom-parser",
        "appendXML",
        vec![frag.clone(), Value::String(Arc::from("<x/>"))],
    );
    let root = call(
        "web:dom-parser",
        "createElement",
        vec![doc, Value::String(Arc::from("root"))],
    );
    call(
        "web:dom-parser",
        "appendChild",
        vec![root.clone(), frag.clone()],
    );
    // The fragment's child moved under root; the fragment is now empty.
    let root_children = prop(&root, "childNodes");
    assert!(matches!(prop(&root_children, "length"), Value::I32(1)));
    let frag_children = prop(&frag, "childNodes");
    assert!(matches!(prop(&frag_children, "length"), Value::I32(0)));
}
