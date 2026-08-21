//! `web:dom-parser` DOM host — `createDocument`
//! (DOM Living Standard §4.5.1 `DOMImplementation.createDocument`).

use std::sync::Arc;
use vybe_compiler::primitives::platforms::register_platforms_all;
use vybe_runtime::{Chunk, Op, VM, Value};

static TEST_GLOBAL_SEQ: std::sync::atomic::AtomicUsize = std::sync::atomic::AtomicUsize::new(0);

fn push_arg(vm: &mut VM, chunk: &mut Chunk, value: Value) {
    match value {
        Value::I32(n) => chunk.emit_i32_const(n, 0),
        Value::I64(n) => chunk.emit_i64_const(n, 0),
        Value::F32(f) => chunk.emit_f32_const(f, 0),
        Value::F64(f) => chunk.emit_f64_const(f, 0),
        Value::Bool(b) => chunk.emit_bool_const(b, 0),
        Value::String(s) => chunk.emit_string_const(&s, 0),
        Value::Null => chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0),
        other => {
            let global = format!(
                "__test_arg_{}",
                TEST_GLOBAL_SEQ.fetch_add(1, std::sync::atomic::Ordering::Relaxed)
            );
            vm.set_global_owned(global.clone(), other);
            let ci = chunk.intern_string_constant(&global);
            chunk.emit_op_u16(Op::GLOBAL_GET, ci, 0);
        }
    }
}

fn call(module: &str, name: &str, args: Vec<Value>) -> Value {
    let mut vm = VM::new();
    register_platforms_all(&mut vm);
    let mut chunk = Chunk::new("<test>");
    let import_idx = chunk.add_import(module, name);
    let argc = args.len() as u8;
    for v in args {
        push_arg(&mut vm, &mut chunk, v);
    }
    chunk.emit_call(import_idx, argc, 0);
    chunk.emit_op(Op::RETURN, 0);
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
