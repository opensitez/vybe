//! WHATWG / W3C web-platform **types** — the runtime TypeRegistry vtables for
//! `TextEncoder`/`TextDecoder`, `URLSearchParams`, fetch `Response`, and the
//! DOM node hierarchy (`Element`/`Document`/`Text`/…).
//!
//! The `register_type` counterpart to the web plugin's host-fn `init`: the web
//! plugin declares its own types here, in its `finalize`. Each method resolves
//! a `web:*` host fn by registry index, so it runs after every plugin's `init`.

use vybe_runtime::{Method, TypeDef};
use vybe_runtime::Framework;

/// Register the web-platform built-in types into the VM's TypeRegistry, and
/// hand the DOM node type ids to `dom_parser` for construction-time stamping.
/// Called from the web plugin's `finalize`.
pub fn register_types(fw: &mut Framework<'_>) {
    // ── TextEncoder ────────────────────────────────────────────────
    {
        let mut t = TypeDef::new("TextEncoder");
        for (method, fname) in &[("encode", "encode"), ("encodeInto", "encodeInto")] {
            if let Some(idx) = fw.host_fn_index("web:encoding", fname) {
                t.methods.insert(method.to_string(), Method::HostFn(idx));
            }
        }
        t.parent = Some(0);
        fw.register_type(t);
    }

    // ── TextDecoder ────────────────────────────────────────────────
    {
        let mut t = TypeDef::new("TextDecoder");
        if let Some(idx) = fw.host_fn_index("web:encoding", "decode") {
            t.methods.insert("decode".to_string(), Method::HostFn(idx));
        }
        t.parent = Some(0);
        fw.register_type(t);
    }

    // ── URLSearchParams ────────────────────────────────────────────
    {
        let mut t = TypeDef::new("URLSearchParams");
        for (method, fname) in &[
            ("get", "searchParamsGet"),
            ("has", "searchParamsHas"),
            ("toString", "searchParamsToString"),
        ] {
            if let Some(idx) = fw.host_fn_index("web:url", fname) {
                t.methods.insert(method.to_string(), Method::HostFn(idx));
            }
        }
        t.parent = Some(0);
        fw.register_type(t);
    }

    // ── Response (fetch result) ────────────────────────────────────
    {
        let mut t = TypeDef::new("Response");
        for (method, fname) in &[("text", "responseText"), ("json", "responseJson")] {
            if let Some(idx) = fw.host_fn_index("web:fetch", fname) {
                t.methods.insert(method.to_string(), Method::HostFn(idx));
            }
        }
        t.parent = Some(0);
        fw.register_type(t);
    }

    // --- WHATWG DOM types (web:dom-parser) -------------------------
    // Method tables for `Document` / `Element` so spec-shaped calls
    // (`elem.querySelector("...")`, `doc.getElementById("x")`,
    // `elem.getAttribute("href")`) dispatch through the TypeRegistry vtable per
    // Component Model resource semantics. Properties (`tagName`, `nodeType`,
    // `childNodes`, …) are set directly on the node Object during parse and
    // resolve via plain `Op::STRUCT_GET` — only the *computed* methods need
    // vtable entries here. `TypeRegistry::resolve_method` lowercases the lookup
    // key, so method-table keys materialise lowercase; the spec-cased names
    // here are for documentation.
    let element_id = {
        let mut t = TypeDef::new("Element");
        for (method, fname) in &[
            // Read API
            ("querySelector", "querySelector"),
            ("querySelectorAll", "querySelectorAll"),
            ("matches", "matches"),
            ("closest", "closest"),
            ("getAttribute", "getAttribute"),
            ("hasAttribute", "hasAttribute"),
            ("getElementsByTagName", "getElementsByTagName"),
            ("getElementsByClassName", "getElementsByClassName"),
            // Mutation API
            ("setAttribute", "setAttribute"),
            ("removeAttribute", "removeAttribute"),
            ("appendChild", "appendChild"),
            ("removeChild", "removeChild"),
            ("insertBefore", "insertBefore"),
            ("replaceChild", "replaceChild"),
            ("cloneNode", "cloneNode"),
            // Namespace-aware (Phase 4)
            ("getAttributeNS", "getAttributeNS"),
            ("hasAttributeNS", "hasAttributeNS"),
            ("setAttributeNS", "setAttributeNS"),
            ("removeAttributeNS", "removeAttributeNS"),
            ("getElementsByTagNameNS", "getElementsByTagNameNS"),
        ] {
            if let Some(idx) = fw.host_fn_index("web:dom-parser", fname) {
                t.methods.insert(method.to_string(), Method::HostFn(idx));
            }
        }
        t.parent = Some(0);
        fw.register_type(t)
    };
    let document_id = {
        let mut t = TypeDef::new("Document");
        for (method, fname) in &[
            // Read API
            ("querySelector", "querySelector"),
            ("querySelectorAll", "querySelectorAll"),
            ("getElementById", "getElementById"),
            ("getElementsByTagName", "getElementsByTagName"),
            ("getElementsByClassName", "getElementsByClassName"),
            // Mutation factories
            ("createElement", "createElement"),
            ("createElementNS", "createElementNS"),
            ("createTextNode", "createTextNode"),
            ("createComment", "createComment"),
            ("createDocumentFragment", "createDocumentFragment"),
            ("appendChild", "appendChild"),
            ("removeChild", "removeChild"),
            ("insertBefore", "insertBefore"),
            ("replaceChild", "replaceChild"),
            ("cloneNode", "cloneNode"),
            ("getElementsByTagNameNS", "getElementsByTagNameNS"),
        ] {
            if let Some(idx) = fw.host_fn_index("web:dom-parser", fname) {
                t.methods.insert(method.to_string(), Method::HostFn(idx));
            }
        }
        t.parent = Some(0);
        fw.register_type(t)
    };
    // Bare placeholder TypeDefs for the other DOM node kinds — no method-table
    // entries today (spec methods like `Text.splitText(offset)` arrive in
    // Phase 3 mutation work).
    let text_id = fw.register_type(TypeDef::new("Text"));
    let comment_id = fw.register_type(TypeDef::new("Comment"));
    let cdata_id = fw.register_type(TypeDef::new("CDATASection"));
    let pi_id = fw.register_type(TypeDef::new("ProcessingInstruction"));
    let attr_id = fw.register_type(TypeDef::new("Attr"));
    let _ = fw.register_type(TypeDef::new("DOMParser"));
    let _ = fw.register_type(TypeDef::new("XMLSerializer"));
    let nnm_id = fw.register_type(TypeDef::new("NamedNodeMap"));

    // Hand the IDs to `dom_parser` so the parser stamps each constructed node's
    // `Object::type_id` for vtable dispatch.
    crate::dom_parser::set_dom_type_ids(crate::dom_parser::DomTypeIds {
        document: document_id,
        element: element_id,
        text: text_id,
        cdata: cdata_id,
        comment: comment_id,
        processing_instruction: pi_id,
        attr: attr_id,
        named_node_map: nnm_id,
    });
}
