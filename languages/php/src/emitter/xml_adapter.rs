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

fn struct_get_key(chunk: &mut Chunk, key: &str, line: u32) {
    let idx = chunk.add_constant(Value::String(Arc::from(key)));
    chunk.emit_op_u16(Op::STRUCT_GET, idx, line);
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
    // → documentElement (the SimpleXML root)
    let chunk = &mut chunks[current];
    struct_get_key(chunk, "documentElement", line);
}
