//! `web:dom-parser` — WHATWG DOM Parsing & Serialization for XML.
//!
//! Implements the spec-shaped surface of `DOMParser`/`XMLSerializer`
//! plus the read side of the DOM Living Standard (Node/Document/Element/
//! Attr/Text/Comment/CDATASection/ProcessingInstruction). Parses with
//! `quick-xml` (pull-based, namespace-aware, full XML 1.0 lexer with
//! CDATA / PI / comment / entity decoding).
//!
//! ## Phases
//! - **Phase 1** (this file, current): read-side DOM —
//!   `parseFromString` / `serializeToString` plus Node/Document/Element
//!   property bag fully populated at parse time so identifier-based
//!   property access (`elem.tagName`, `elem.children[0].textContent`)
//!   works through `Op::STRUCT_GET`. Helper host fns provide the
//!   computed methods that walk the tree (`getElementById`,
//!   `getElementsByTagName`, `getAttribute`).
//! - **Phase 2** (planned): Selectors API (`querySelector`,
//!   `querySelectorAll`) — CSS selector subset matcher.
//! - **Phase 3** (planned): mutation —
//!   `createElement`/`setAttribute`/`appendChild`/`removeChild`/
//!   `insertBefore` plus `Document.create*` factories.
//! - **Phase 4** (planned): namespaces (`*NS` variants),
//!   `cloneNode(deep)`, `DocumentFragment`, `NodeFilter`/`TreeWalker`.
//!
//! ## Node-shape contract
//! Every DOM node is a Vybe `Object` with:
//!
//!   * `__type` — `"Document"` / `"Element"` / `"Text"` / `"Comment"` /
//!     `"CDATASection"` / `"ProcessingInstruction"` / `"Attr"`.
//!   * `nodeType` — WHATWG number (1=Element, 3=Text, 4=CDATA, 7=PI,
//!     8=Comment, 9=Document).
//!   * `nodeName` — `"div"` for elements, `"#text"` for text, `"#comment"`
//!     for comments, `"#cdata-section"` for CDATA, `"#document"` for
//!     the Document, the target for PIs.
//!   * `nodeValue` — text payload for Text/Comment/CDATA/PI/Attr,
//!     `null` for Element/Document.
//!   * `textContent` — concatenated descendant text (computed eagerly).
//!   * `childNodes` — Array of all child nodes.
//!   * `parentNode` — back-ref to parent (set during post-walk).
//!   * `ownerDocument` — back-ref to the Document.
//!   * `firstChild` / `lastChild` / `nextSibling` / `previousSibling` —
//!     navigation helpers wired post-walk.
//!
//! Element nodes additionally carry:
//!   * `tagName` (mirror of `nodeName` for elements)
//!   * `localName`, `prefix`, `namespaceURI` (Phase 4 — placeholder `null` today)
//!   * `attributes` — Object (NamedNodeMap-shape) keyed by attr name
//!   * `children` — Array of Element-only children (HTMLCollection-shape)
//!   * `firstElementChild` / `lastElementChild`
//!   * `id`, `className` (mirrors of common attributes)
//!
//! Document nodes additionally carry:
//!   * `documentElement` — the root Element
//!   * `doctype` (Phase 4 — `null` today)

use std::sync::{Arc, Mutex, OnceLock};
use vybe_runtime::value::{Object, ObjectKind};
use vybe_runtime::{VM, Value};

use quick_xml::Reader;
use quick_xml::events::Event;

// The DOM operations the HTML sink builds through. Only the seam — this module
// names no engine, which is what keeps the tree it builds swappable.
use crate::engine::{DomOp, DomValue};

// ── nodeType constants (WHATWG DOM Living Standard §4.4) ──────────
const ELEMENT_NODE: i32 = 1;
const TEXT_NODE: i32 = 3;
const CDATA_SECTION_NODE: i32 = 4;
const PROCESSING_INSTRUCTION_NODE: i32 = 7;
const COMMENT_NODE: i32 = 8;
const DOCUMENT_NODE: i32 = 9;
const DOCUMENT_FRAGMENT_NODE: i32 = 11;

// ── DOM type-id table ─────────────────────────────────────────────
// Captured by `register_dom_type_ids` once `builtin_types::register_all`
// runs. Host fns (which can't see the type registry at registration
// time) read these IDs at parse time so each materialised node carries
// the correct `Object::type_id` for TypeRegistry vtable dispatch
// (`elem.querySelector(...)`, etc.).

#[derive(Default, Clone, Copy, Debug)]
pub struct DomTypeIds {
    pub document: usize,
    pub element: usize,
    pub text: usize,
    pub cdata: usize,
    pub comment: usize,
    pub processing_instruction: usize,
    pub attr: usize,
    pub named_node_map: usize,
}

static DOM_TYPE_IDS: OnceLock<DomTypeIds> = OnceLock::new();

/// Called from `builtin_types::register_all` once the TypeRegistry has
/// been populated with `Element` / `Document` / etc. After this the
/// parser will stamp `Object::type_id` so method dispatch through
/// `resolve_property` finds the registered vtable.
pub fn set_dom_type_ids(ids: DomTypeIds) {
    let _ = DOM_TYPE_IDS.set(ids);
}

fn dom_type_ids() -> Option<&'static DomTypeIds> {
    DOM_TYPE_IDS.get()
}

fn type_id_for(node_type: i32) -> usize {
    let Some(ids) = dom_type_ids() else { return 0 };
    match node_type {
        DOCUMENT_NODE => ids.document,
        ELEMENT_NODE => ids.element,
        TEXT_NODE => ids.text,
        CDATA_SECTION_NODE => ids.cdata,
        COMMENT_NODE => ids.comment,
        PROCESSING_INSTRUCTION_NODE => ids.processing_instruction,
        _ => 0,
    }
}

// ── Public registration ───────────────────────────────────────────

pub fn register(vm: &mut VM) {
    // ── DOMParser ────────────────────────────────────────────────
    vm.register_host_fn(
        "web:dom-parser",
        "parserNew",
        Box::new(|_ctx, _args| {
            let mut obj = Object::new();
            obj.properties.insert("__type".into(), s("DOMParser"));
            Value::Object(vybe_runtime::heap::alloc(obj))
        }),
    );

    vm.register_host_fn(
        "web:dom-parser",
        "parseFromString",
        Box::new(|_ctx, args| {
            // Spec: `DOMParser.parseFromString(string, type)`. We accept
            // both the instance-call shape (args[0] = DOMParser, args[1]
            // = string, args[2] = type) and the flat shorthand
            // (args[0] = string, args[1] = type) so callers that don't
            // construct a parser still work.
            let (xml_arg, type_arg) = match args.first() {
                Some(Value::Object(o))
                    if o.lock()
                        .unwrap()
                        .properties
                        .get("__type")
                        .and_then(|v| match v {
                            Value::String(s) => Some(s.as_ref().to_string()),
                            _ => None,
                        })
                        .as_deref()
                        == Some("DOMParser") =>
                {
                    (
                        args.get(1).cloned().unwrap_or(Value::Null),
                        args.get(2).cloned().unwrap_or(Value::Null),
                    )
                }
                _ => (
                    args.first().cloned().unwrap_or(Value::Null),
                    args.get(1).cloned().unwrap_or(Value::Null),
                ),
            };
            let xml = match xml_arg {
                Value::String(s) => s.to_string(),
                other => format!("{}", other),
            };
            // The type argument decides the GRAMMAR, which is the whole
            // reason the spec has one. `text/html` reads HTML; every other
            // type — and an absent one — reads XML, which is what the
            // `parse` shorthand and every existing caller already get.
            // Discarding this argument is why a bare `<br>` used to produce
            // a `<parsererror>` document.
            //
            // It also decides which DOM answers. `text/html` builds a real
            // document — the tree `window.document` is, which cascades, lays
            // out and paints; every other type builds the property-bag tree
            // that `DOMDocument`, XLinq and `xml2js` read. The spec returns two
            // kinds of document here too (`HTMLDocument` and `XMLDocument`);
            // what is still ours to close is that they come from two engines
            // rather than one. See `TreeSink`.
            let html = matches!(&type_arg, Value::String(t)
                if t.trim().eq_ignore_ascii_case("text/html"));
            if html {
                return parse_html_document(&xml);
            }
            match parse_markup(&xml, Grammar::Xml) {
                Ok(doc) => doc,
                Err(_) => parse_error_document(&xml),
            }
        }),
    );

    // Convenience flat fn: `parse(s)` shorthand callers use without an
    // explicit DOMParser instance. Same return as `parseFromString`.
    vm.register_host_fn(
        "web:dom-parser",
        "parse",
        Box::new(|_ctx, args| {
            let xml = match args.first() {
                Some(Value::String(s)) => s.to_string(),
                Some(other) => format!("{}", other),
                None => String::new(),
            };
            match parse_xml(&xml) {
                Ok(doc) => doc,
                Err(_) => parse_error_document(&xml),
            }
        }),
    );

    // ── XMLSerializer ────────────────────────────────────────────
    vm.register_host_fn(
        "web:dom-parser",
        "serializerNew",
        Box::new(|_ctx, _args| {
            let mut obj = Object::new();
            obj.properties.insert("__type".into(), s("XMLSerializer"));
            Value::Object(vybe_runtime::heap::alloc(obj))
        }),
    );

    vm.register_host_fn(
        "web:dom-parser",
        "serializeToString",
        Box::new(|_ctx, args| {
            // Accepts both `(serializer, node)` and `(node)` shapes.
            let node = match args.first() {
                Some(Value::Object(o))
                    if o.lock()
                        .unwrap()
                        .properties
                        .get("__type")
                        .and_then(|v| match v {
                            Value::String(s) => Some(s.as_ref().to_string()),
                            _ => None,
                        })
                        .as_deref()
                        == Some("XMLSerializer") =>
                {
                    args.get(1).cloned().unwrap_or(Value::Null)
                }
                other => other.cloned().unwrap_or(Value::Null),
            };
            let mut out = String::new();
            if let Value::Object(o) = &node {
                serialize_node(o, &mut out);
            }
            s(&out)
        }),
    );

    // Convenience: `toString(node)` matches the legacy shorthand.
    vm.register_host_fn(
        "web:dom-parser",
        "toString",
        Box::new(|_ctx, args| {
            let mut out = String::new();
            if let Some(Value::Object(o)) = args.first() {
                serialize_node(o, &mut out);
            }
            s(&out)
        }),
    );

    // Convenience helper retained from the pre-spec API: `load(path)`
    // reads from disk and parses. Real WHATWG flow is `fetch(url)` →
    // `.text()` → `parseFromString` — provided here for parity with
    // VB `XDocument.Load(path)` test patterns.
    vm.register_host_fn(
        "web:dom-parser",
        "load",
        Box::new(|_ctx, args| {
            let path = match args.first() {
                Some(Value::String(s)) => s.to_string(),
                Some(other) => format!("{}", other),
                None => return Value::Null,
            };
            match std::fs::read_to_string(&path) {
                Ok(xml) => parse_xml(&xml).unwrap_or_else(|_| parse_error_document(&xml)),
                Err(_) => Value::Null,
            }
        }),
    );

    // ── Node / Element / Document method helpers ─────────────────
    // Properties (`tagName`, `nodeType`, `childNodes`, etc.) are set
    // directly on the object during parse, so user code reads them
    // via plain `Op::STRUCT_GET` — no host fn round-trip. The host
    // fns below cover the *computed* methods that walk the tree.

    vm.register_host_fn(
        "web:dom-parser",
        "getElementById",
        Box::new(|_ctx, args| {
            let Some(Value::Object(root)) = args.first() else {
                return Value::Null;
            };
            let target = match args.get(1) {
                Some(Value::String(s)) => s.to_string(),
                Some(other) => format!("{}", other),
                None => return Value::Null,
            };
            find_by_id(root, &target).unwrap_or(Value::Null)
        }),
    );

    vm.register_host_fn(
        "web:dom-parser",
        "getElementsByTagName",
        Box::new(|_ctx, args| {
            let Some(Value::Object(root)) = args.first() else {
                return empty_array();
            };
            let tag = match args.get(1) {
                Some(Value::String(s)) => s.to_string(),
                Some(other) => format!("{}", other),
                None => return empty_array(),
            };
            let mut out: Vec<Value> = Vec::new();
            find_by_tag_name(root, &tag, &mut out);
            make_array(out)
        }),
    );

    vm.register_host_fn(
        "web:dom-parser",
        "getElementsByClassName",
        Box::new(|_ctx, args| {
            let Some(Value::Object(root)) = args.first() else {
                return empty_array();
            };
            let cls = match args.get(1) {
                Some(Value::String(s)) => s.to_string(),
                Some(other) => format!("{}", other),
                None => return empty_array(),
            };
            let mut out: Vec<Value> = Vec::new();
            find_by_class_name(root, &cls, &mut out);
            make_array(out)
        }),
    );

    vm.register_host_fn(
        "web:dom-parser",
        "getAttribute",
        Box::new(|_ctx, args| {
            let Some(Value::Object(elem)) = args.first() else {
                return Value::Null;
            };
            let name = match args.get(1) {
                Some(Value::String(s)) => s.to_string(),
                Some(other) => format!("{}", other),
                None => return Value::Null,
            };
            let elem = elem.lock().unwrap();
            let Some(Value::Object(attrs)) = elem.properties.get("attributes") else {
                return Value::Null;
            };
            let attrs = attrs.lock().unwrap();
            attrs.properties.get(&name).cloned().unwrap_or(Value::Null)
        }),
    );

    vm.register_host_fn(
        "web:dom-parser",
        "hasAttribute",
        Box::new(|_ctx, args| {
            let Some(Value::Object(elem)) = args.first() else {
                return Value::Bool(false);
            };
            let name = match args.get(1) {
                Some(Value::String(s)) => s.to_string(),
                Some(other) => format!("{}", other),
                None => return Value::Bool(false),
            };
            let elem = elem.lock().unwrap();
            let Some(Value::Object(attrs)) = elem.properties.get("attributes") else {
                return Value::Bool(false);
            };
            Value::Bool(attrs.lock().unwrap().properties.contains_key(&name))
        }),
    );

    // ── Selectors API Level 1 (W3C) ──────────────────────────────
    // CSS selector subset: tag, `*`, `#id`, `.class`, `[attr]` /
    // `[attr="val"]` / `[attr*="val"]` / `[attr^="val"]` /
    // `[attr$="val"]`, descendant (` `), child (`>`), adjacent sibling
    // (`+`), general sibling (`~`), compound (`tag.class#id[attr]`),
    // and selector lists (comma-separated). Pseudo-classes
    // (`:first-child`, `:nth-child`) are Phase 2.5 follow-up.

    vm.register_host_fn(
        "web:dom-parser",
        "querySelector",
        Box::new(|_ctx, args| {
            let Some(Value::Object(root)) = args.first() else {
                return Value::Null;
            };
            let selector = match args.get(1) {
                Some(Value::String(s)) => s.to_string(),
                Some(other) => format!("{}", other),
                None => return Value::Null,
            };
            let parsed = match parse_selector_list(&selector) {
                Some(s) => s,
                None => return Value::Null,
            };
            find_first_match(root, &parsed).unwrap_or(Value::Null)
        }),
    );

    vm.register_host_fn(
        "web:dom-parser",
        "querySelectorAll",
        Box::new(|_ctx, args| {
            let Some(Value::Object(root)) = args.first() else {
                return empty_array();
            };
            let selector = match args.get(1) {
                Some(Value::String(s)) => s.to_string(),
                Some(other) => format!("{}", other),
                None => return empty_array(),
            };
            let parsed = match parse_selector_list(&selector) {
                Some(s) => s,
                None => return empty_array(),
            };
            let mut out: Vec<Value> = Vec::new();
            find_all_matches(root, &parsed, &mut out);
            make_array(out)
        }),
    );

    vm.register_host_fn(
        "web:dom-parser",
        "matches",
        Box::new(|_ctx, args| {
            let Some(Value::Object(elem)) = args.first() else {
                return Value::Bool(false);
            };
            let selector = match args.get(1) {
                Some(Value::String(s)) => s.to_string(),
                Some(other) => format!("{}", other),
                None => return Value::Bool(false),
            };
            let parsed = match parse_selector_list(&selector) {
                Some(s) => s,
                None => return Value::Bool(false),
            };
            Value::Bool(parsed.iter().any(|sel| selector_matches(elem, sel)))
        }),
    );

    // ── DOM Mutation (Phase 3) ───────────────────────────────────
    // WHATWG DOM Living Standard §4.5: Document factories +
    // Node mutation methods. Newly-created nodes start orphaned
    // (parentNode=null, no siblings, ownerDocument set when known).

    vm.register_host_fn("web:dom-parser", "createElement", Box::new(|_ctx, args| {
        let tag = match args.get(1).or(args.first()) {
            Some(Value::String(s)) => s.to_string(),
            Some(other) => format!("{}", other),
            None => return Value::Null };
        // Detect (doc, tag) shape vs (tag) shape: if first arg is a
        // Document, owner is set; otherwise orphan.
        let owner: Option<Arc<Mutex<Object>>> = match args.first() {
            Some(Value::Object(o)) => {
                let lock = o.lock().unwrap();
                let is_doc = matches!(lock.properties.get("nodeType"), Some(Value::I32(t)) if *t == DOCUMENT_NODE);
                if is_doc { Some(o.clone()) } else { None }
            }
            _ => None };
        new_element_node(&tag, owner.as_ref())
    }));

    // DOMImplementation.createDocument(namespace, qualifiedName, doctype)
    // — DOM Living Standard §4.5.1 "createDocument". Returns a new document;
    // when `qualifiedName` is a non-empty string a document element is created
    // (in `namespace`) and appended as the root. `new Document()` maps here
    // with no arguments (empty document).
    vm.register_host_fn(
        "web:dom-parser",
        "createDocument",
        Box::new(|_ctx, args| {
            let namespace = match args.first() {
                Some(Value::String(ns)) if !ns.is_empty() => Some(ns.to_string()),
                _ => None,
            };
            let qualified = match args.get(1) {
                Some(Value::String(q)) if !q.is_empty() => Some(q.to_string()),
                _ => None,
            };
            let doc = make_node(DOCUMENT_NODE, "#document", None);
            if let (Value::Object(d), Some(name)) = (&doc, qualified.as_ref()) {
                let elem = new_element_node(name, Some(d));
                if let (Value::Object(e), Some(ns)) = (&elem, namespace.as_ref()) {
                    e.lock()
                        .unwrap()
                        .properties
                        .insert("namespaceURI".into(), s(ns));
                }
                if let Value::Object(child) = &elem {
                    append_child_inner(d, child);
                    d.lock()
                        .unwrap()
                        .properties
                        .insert("documentElement".into(), elem.clone());
                }
            }
            doc
        }),
    );

    vm.register_host_fn("web:dom-parser", "createTextNode", Box::new(|_ctx, args| {
        let text = match args.get(1).or(args.first()) {
            Some(Value::String(s)) => s.to_string(),
            Some(other) => format!("{}", other),
            None => String::new() };
        let owner: Option<Arc<Mutex<Object>>> = match args.first() {
            Some(Value::Object(o)) => {
                let lock = o.lock().unwrap();
                let is_doc = matches!(lock.properties.get("nodeType"), Some(Value::I32(t)) if *t == DOCUMENT_NODE);
                if is_doc { Some(o.clone()) } else { None }
            }
            _ => None };
        let node = make_node(TEXT_NODE, "#text", Some(&text));
        if let (Value::Object(n), Some(d)) = (&node, &owner) {
            n.lock().unwrap().properties.insert("ownerDocument".into(), Value::Object(d.clone()));
        }
        if let Value::Object(n) = &node {
            n.lock().unwrap().properties.insert("textContent".into(), s(&text));
        }
        node
    }));

    // Document.createCDATASection(data) — DOM Living Standard §4.5 (Document
    // interface). Produces a CDATASection node (nodeType 4, nodeName
    // "#cdata-section") whose text is `data`; `ownerDocument` links back when
    // called on a document.
    vm.register_host_fn(
        "web:dom-parser",
        "createCDATASection",
        Box::new(|_ctx, args| {
            let text = match args.get(1).or(args.first()) {
                Some(Value::String(s)) => s.to_string(),
                Some(other) => format!("{}", other),
                None => String::new() };
            let owner: Option<Arc<Mutex<Object>>> = match args.first() {
                Some(Value::Object(o)) => {
                    let lock = o.lock().unwrap();
                    let is_doc = matches!(lock.properties.get("nodeType"), Some(Value::I32(t)) if *t == DOCUMENT_NODE);
                    if is_doc { Some(o.clone()) } else { None }
                }
                _ => None };
            let node = make_node(CDATA_SECTION_NODE, "#cdata-section", Some(&text));
            if let Value::Object(n) = &node {
                let mut nl = n.lock().unwrap();
                nl.properties.insert("textContent".into(), s(&text));
                if let Some(d) = &owner {
                    nl.properties
                        .insert("ownerDocument".into(), Value::Object(d.clone()));
                }
            }
            node
        }),
    );

    vm.register_host_fn(
        "web:dom-parser",
        "createComment",
        Box::new(|_ctx, args| {
            let text = match args.get(1).or(args.first()) {
                Some(Value::String(s)) => s.to_string(),
                Some(other) => format!("{}", other),
                None => String::new(),
            };
            let node = make_node(COMMENT_NODE, "#comment", Some(&text));
            if let Value::Object(n) = &node {
                n.lock()
                    .unwrap()
                    .properties
                    .insert("textContent".into(), s(&text));
            }
            node
        }),
    );

    vm.register_host_fn(
        "web:dom-parser",
        "setAttribute",
        Box::new(|_ctx, args| {
            let Some(Value::Object(elem)) = args.first() else {
                return Value::Null;
            };
            let name = match args.get(1) {
                Some(Value::String(s)) => s.to_string(),
                Some(other) => format!("{}", other),
                None => return Value::Null,
            };
            let val = match args.get(2) {
                Some(Value::String(s)) => s.to_string(),
                Some(other) => format!("{}", other),
                None => String::new(),
            };
            let elem_lock = elem.lock().unwrap();
            let attrs = elem_lock.properties.get("attributes").cloned();
            drop(elem_lock);
            if let Some(Value::Object(a)) = attrs {
                a.lock().unwrap().properties.insert(name.clone(), s(&val));
            }
            // Mirror id / className for the common attrs.
            let mut elem_w = elem.lock().unwrap();
            if name == "id" {
                elem_w.properties.insert("id".into(), s(&val));
            } else if name == "class" {
                elem_w.properties.insert("className".into(), s(&val));
            }
            Value::Null
        }),
    );

    vm.register_host_fn(
        "web:dom-parser",
        "removeAttribute",
        Box::new(|_ctx, args| {
            let Some(Value::Object(elem)) = args.first() else {
                return Value::Null;
            };
            let name = match args.get(1) {
                Some(Value::String(s)) => s.to_string(),
                Some(other) => format!("{}", other),
                None => return Value::Null,
            };
            let elem_lock = elem.lock().unwrap();
            let attrs = elem_lock.properties.get("attributes").cloned();
            drop(elem_lock);
            if let Some(Value::Object(a)) = attrs {
                a.lock().unwrap().properties.shift_remove(&name);
            }
            let mut elem_w = elem.lock().unwrap();
            if name == "id" {
                elem_w.properties.insert("id".into(), s(""));
            } else if name == "class" {
                elem_w.properties.insert("className".into(), s(""));
            }
            Value::Null
        }),
    );

    vm.register_host_fn(
        "web:dom-parser",
        "appendChild",
        Box::new(|_ctx, args| {
            let Some(Value::Object(parent)) = args.first() else {
                return Value::Null;
            };
            let Some(Value::Object(child)) = args.get(1) else {
                return Value::Null;
            };
            // DOM §4.2.1 pre-insert: appending a DocumentFragment moves ITS
            // children into the parent, not the fragment node itself.
            let is_fragment = matches!(
                child.lock().unwrap().properties.get("nodeType"),
                Some(Value::I32(t)) if *t == DOCUMENT_FRAGMENT_NODE
            );
            if is_fragment {
                let moved: Vec<Value> = {
                    let c = child.lock().unwrap();
                    match c.properties.get("childNodes") {
                        Some(Value::Object(arr)) => match &arr.lock().unwrap().kind {
                            ObjectKind::Array(items) => items.clone(),
                            _ => Vec::new(),
                        },
                        _ => Vec::new(),
                    }
                };
                // Empty the fragment, then append each former child to parent.
                if let Some(Value::Object(arr)) = child.lock().unwrap().properties.get("childNodes")
                {
                    if let ObjectKind::Array(ref mut items) = arr.lock().unwrap().kind {
                        items.clear();
                    }
                }
                refresh_node_relationships(child);
                for node in &moved {
                    if let Value::Object(n) = node {
                        append_child_inner(parent, n);
                    }
                }
                return Value::Object(child.clone());
            }
            // Detach child from any current parent first (DOM semantics:
            // appendChild is a move, not a copy — `pre-insert` step in spec).
            detach_from_parent(child);
            // Append + rewire siblings.
            append_child_inner(parent, child);
            Value::Object(child.clone())
        }),
    );

    vm.register_host_fn(
        "web:dom-parser",
        "removeChild",
        Box::new(|_ctx, args| {
            let Some(Value::Object(parent)) = args.first() else {
                return Value::Null;
            };
            let Some(Value::Object(child)) = args.get(1) else {
                return Value::Null;
            };
            if remove_child_inner(parent, child) {
                Value::Object(child.clone())
            } else {
                Value::Null
            }
        }),
    );

    vm.register_host_fn(
        "web:dom-parser",
        "insertBefore",
        Box::new(|_ctx, args| {
            let Some(Value::Object(parent)) = args.first() else {
                return Value::Null;
            };
            let Some(Value::Object(new_child)) = args.get(1) else {
                return Value::Null;
            };
            let reference = args.get(2).cloned();
            detach_from_parent(new_child);
            match reference {
                Some(Value::Object(ref_node)) => {
                    if !insert_before_inner(parent, new_child, &ref_node) {
                        append_child_inner(parent, new_child);
                    }
                }
                _ => append_child_inner(parent, new_child),
            }
            Value::Object(new_child.clone())
        }),
    );

    vm.register_host_fn(
        "web:dom-parser",
        "replaceChild",
        Box::new(|_ctx, args| {
            let Some(Value::Object(parent)) = args.first() else {
                return Value::Null;
            };
            let Some(Value::Object(new_child)) = args.get(1) else {
                return Value::Null;
            };
            let Some(Value::Object(old_child)) = args.get(2) else {
                return Value::Null;
            };
            detach_from_parent(new_child);
            if !insert_before_inner(parent, new_child, old_child) {
                return Value::Null;
            }
            let _ = remove_child_inner(parent, old_child);
            Value::Object(old_child.clone())
        }),
    );

    vm.register_host_fn(
        "web:dom-parser",
        "cloneNode",
        Box::new(|_ctx, args| {
            let Some(Value::Object(node)) = args.first() else {
                return Value::Null;
            };
            let deep = matches!(args.get(1), Some(Value::Bool(true)));
            clone_node(node, deep)
        }),
    );

    // ── DocumentFragment + Namespace-aware variants (Phase 4) ────
    vm.register_host_fn("web:dom-parser", "createDocumentFragment", Box::new(|_ctx, args| {
        let owner: Option<Arc<Mutex<Object>>> = match args.first() {
            Some(Value::Object(o)) => {
                let lock = o.lock().unwrap();
                let is_doc = matches!(lock.properties.get("nodeType"), Some(Value::I32(t)) if *t == DOCUMENT_NODE);
                if is_doc { Some(o.clone()) } else { None }
            }
            _ => None };
        new_document_fragment(owner.as_ref())
    }));

    // PHP `DOMDocumentFragment::appendXML($xml)` — parse the XML fragment and
    // append the resulting top-level nodes to the fragment. Not part of the
    // DOM standard, but composed here entirely from the spec parse surface.
    vm.register_host_fn(
        "web:dom-parser",
        "appendXML",
        Box::new(|_ctx, args| {
            let Some(Value::Object(frag)) = args.first() else {
                return Value::Bool(false);
            };
            let xml = match args.get(1) {
                Some(Value::String(s)) => s.to_string(),
                Some(other) => format!("{}", other),
                None => return Value::Bool(false),
            };
            let doc = match parse_xml(&xml) {
                Ok(doc) => doc,
                Err(_) => return Value::Bool(false),
            };
            // Move the parsed document's top-level nodes into the fragment.
            let roots: Vec<Value> = if let Value::Object(d) = &doc {
                match d.lock().unwrap().properties.get("childNodes") {
                    Some(Value::Object(arr)) => match &arr.lock().unwrap().kind {
                        ObjectKind::Array(items) => items.clone(),
                        _ => Vec::new(),
                    },
                    _ => Vec::new(),
                }
            } else {
                Vec::new()
            };
            for node in &roots {
                if let Value::Object(n) = node {
                    append_child_inner(frag, n);
                }
            }
            Value::Bool(true)
        }),
    );

    vm.register_host_fn("web:dom-parser", "createElementNS", Box::new(|_ctx, args| {
        let owner: Option<Arc<Mutex<Object>>> = match args.first() {
            Some(Value::Object(o)) => {
                let lock = o.lock().unwrap();
                let is_doc = matches!(lock.properties.get("nodeType"), Some(Value::I32(t)) if *t == DOCUMENT_NODE);
                if is_doc { Some(o.clone()) } else { None }
            }
            _ => None };
        let (ns_arg, qname_arg) = if owner.is_some() {
            (args.get(1), args.get(2))
        } else {
            (args.first(), args.get(1))
        };
        let ns = match ns_arg {
            Some(Value::String(s)) => Some(s.to_string()),
            Some(Value::Null) | None => None,
            Some(other) => Some(format!("{}", other)) };
        let qname = match qname_arg {
            Some(Value::String(s)) => s.to_string(),
            Some(other) => format!("{}", other),
            None => return Value::Null };
        let elem = new_element_node(&qname, owner.as_ref());
        if let Value::Object(e) = &elem {
            e.lock().unwrap().properties.insert("namespaceURI".into(),
                ns.map(|n| s(&n)).unwrap_or(Value::Null));
        }
        elem
    }));

    vm.register_host_fn(
        "web:dom-parser",
        "setAttributeNS",
        Box::new(|_ctx, args| {
            let Some(Value::Object(elem)) = args.first() else {
                return Value::Null;
            };
            let _ns = match args.get(1) {
                Some(Value::String(s)) => Some(s.to_string()),
                Some(Value::Null) | None => None,
                Some(other) => Some(format!("{}", other)),
            };
            let name = match args.get(2) {
                Some(Value::String(s)) => s.to_string(),
                Some(other) => format!("{}", other),
                None => return Value::Null,
            };
            let val = match args.get(3) {
                Some(Value::String(s)) => s.to_string(),
                Some(other) => format!("{}", other),
                None => String::new(),
            };
            let attrs = elem.lock().unwrap().properties.get("attributes").cloned();
            if let Some(Value::Object(a)) = attrs {
                a.lock().unwrap().properties.insert(name.clone(), s(&val));
            }
            if name == "id" {
                elem.lock().unwrap().properties.insert("id".into(), s(&val));
            } else if name == "class" {
                elem.lock()
                    .unwrap()
                    .properties
                    .insert("className".into(), s(&val));
            }
            Value::Null
        }),
    );

    vm.register_host_fn(
        "web:dom-parser",
        "getAttributeNS",
        Box::new(|_ctx, args| {
            let Some(Value::Object(elem)) = args.first() else {
                return Value::Null;
            };
            let name = match args.get(2).or(args.get(1)) {
                Some(Value::String(s)) => s.to_string(),
                Some(other) => format!("{}", other),
                None => return Value::Null,
            };
            let attrs = elem.lock().unwrap().properties.get("attributes").cloned();
            if let Some(Value::Object(a)) = attrs {
                return a
                    .lock()
                    .unwrap()
                    .properties
                    .get(&name)
                    .cloned()
                    .unwrap_or(Value::Null);
            }
            Value::Null
        }),
    );

    vm.register_host_fn(
        "web:dom-parser",
        "hasAttributeNS",
        Box::new(|_ctx, args| {
            let Some(Value::Object(elem)) = args.first() else {
                return Value::Bool(false);
            };
            let name = match args.get(2).or(args.get(1)) {
                Some(Value::String(s)) => s.to_string(),
                Some(other) => format!("{}", other),
                None => return Value::Bool(false),
            };
            let attrs = elem.lock().unwrap().properties.get("attributes").cloned();
            if let Some(Value::Object(a)) = attrs {
                return Value::Bool(a.lock().unwrap().properties.contains_key(&name));
            }
            Value::Bool(false)
        }),
    );

    vm.register_host_fn(
        "web:dom-parser",
        "removeAttributeNS",
        Box::new(|_ctx, args| {
            let Some(Value::Object(elem)) = args.first() else {
                return Value::Null;
            };
            let name = match args.get(2).or(args.get(1)) {
                Some(Value::String(s)) => s.to_string(),
                Some(other) => format!("{}", other),
                None => return Value::Null,
            };
            let attrs = elem.lock().unwrap().properties.get("attributes").cloned();
            if let Some(Value::Object(a)) = attrs {
                a.lock().unwrap().properties.shift_remove(&name);
            }
            Value::Null
        }),
    );

    vm.register_host_fn(
        "web:dom-parser",
        "getElementsByTagNameNS",
        Box::new(|_ctx, args| {
            let Some(Value::Object(root)) = args.first() else {
                return empty_array();
            };
            let local = match args.get(2).or(args.get(1)) {
                Some(Value::String(s)) => s.to_string(),
                Some(other) => format!("{}", other),
                None => return empty_array(),
            };
            let mut out: Vec<Value> = Vec::new();
            find_by_local_name(root, &local, &mut out);
            make_array(out)
        }),
    );

    vm.register_host_fn(
        "web:dom-parser",
        "closest",
        Box::new(|_ctx, args| {
            let Some(Value::Object(elem)) = args.first() else {
                return Value::Null;
            };
            let selector = match args.get(1) {
                Some(Value::String(s)) => s.to_string(),
                Some(other) => format!("{}", other),
                None => return Value::Null,
            };
            let parsed = match parse_selector_list(&selector) {
                Some(s) => s,
                None => return Value::Null,
            };
            let mut current = Some(elem.clone());
            while let Some(node) = current {
                if parsed.iter().any(|sel| selector_matches(&node, sel)) {
                    return Value::Object(node);
                }
                let parent = {
                    let n = node.lock().unwrap();
                    match n.properties.get("parentNode") {
                        Some(Value::Object(p)) => Some(p.clone()),
                        _ => None,
                    }
                };
                current = parent;
            }
            Value::Null
        }),
    );
}

// ── Parser entry point ────────────────────────────────────────────

// ── HTML tolerance ────────────────────────────────────────────────
//
// `parseFromString(s, "text/html")` used to discard its type argument and
// read HTML with an XML lexer. That is not a stricter reading of HTML — it
// is the wrong grammar: a bare `<br>` is well-formed HTML and a fatal XML
// error, so an ordinary page produced a `<parsererror>` document and
// nothing else. What follows is the part of HTML that XML does not have.
//
// Scope, stated so it is a boundary rather than an oversight: the
// tokenizer's void-element rule, the in-body insertion mode's
// implied-end-tag rules, tag/attribute name folding, tolerated end tags,
// and the common named character references. NOT the full tree
// construction algorithm — no adoption agency, no foster parenting, no
// `<head>`/`<body>` auto-insertion, no template contents. Markup needing
// those parses without error and produces a tree a browser would nest
// differently. `__parseRecoveries` below is how that stays visible.

/// Which grammar the reader is reading.
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
enum Grammar {
    Xml,
    Html,
}

/// Elements with no content and no end tag.
///
/// <https://html.spec.whatwg.org/multipage/syntax.html#void-elements>
///
/// This is the single rule that decides whether real-world HTML parses at
/// all: without it `<br>` opens an element that is never closed and every
/// following sibling nests one level deeper.
const VOID_ELEMENTS: &[&str] = &[
    "area", "base", "br", "col", "embed", "hr", "img", "input", "link", "meta", "source", "track",
    "wbr",
];

/// Elements whose content is raw text, not markup.
///
/// <https://html.spec.whatwg.org/multipage/syntax.html#raw-text-elements>
///
/// Inside these, `<` and `&` are ordinary characters. Read as markup,
/// `if (a < b)` opens an element named `b)` and the rest of the script
/// disappears into it — measured, and it happened SILENTLY: no error, no
/// recovery, just a script whose text stopped at the comparison.
const RAW_TEXT_ELEMENTS: &[&str] = &["script", "style"];

/// Block-level start tags that close an open `<p>`.
///
/// <https://html.spec.whatwg.org/multipage/syntax.html#optional-tags>
const CLOSES_A_PARAGRAPH: &[&str] = &[
    "address",
    "article",
    "aside",
    "blockquote",
    "details",
    "div",
    "dl",
    "fieldset",
    "figcaption",
    "figure",
    "footer",
    "form",
    "h1",
    "h2",
    "h3",
    "h4",
    "h5",
    "h6",
    "header",
    "hgroup",
    "hr",
    "main",
    "menu",
    "nav",
    "ol",
    "p",
    "pre",
    "section",
    "table",
    "ul",
];

/// Does starting `opening` imply the end of an open `open` element?
///
/// The omitted-end-tag rules, which is why `<li>a<li>b` is two siblings in
/// a browser and was two NESTED items under an XML reading.
fn implies_end_of(open: &str, opening: &str) -> bool {
    match open {
        "p" => CLOSES_A_PARAGRAPH.contains(&opening),
        "li" => opening == "li",
        "dt" | "dd" => opening == "dt" || opening == "dd",
        "option" => opening == "option" || opening == "optgroup",
        "optgroup" => opening == "optgroup",
        "tr" => opening == "tr",
        "td" | "th" => opening == "td" || opening == "th" || opening == "tr",
        "thead" | "tbody" => opening == "tbody" || opening == "tfoot",
        _ => false,
    }
}

/// The named character references worth carrying.
///
/// HTML defines about 2200 and they are a generated table; this is the set
/// that appears in ordinary prose. An unlisted name is left as written
/// rather than dropped — `&fnof;` reads back as `&fnof;`, which is wrong
/// but visible, where an empty string would silently delete text.
fn html_entity(name: &str) -> Option<&'static str> {
    Some(match name {
        // The five XML predefines itself. `unescape_with` does NOT keep them
        // — it routes every named reference to the resolver — so leaving
        // them out made `&amp;` read back as the literal `&amp;` in HTML
        // while still decoding in XML.
        "amp" => "&",
        "lt" => "<",
        "gt" => ">",
        "quot" => "\"",
        "apos" => "'",
        "nbsp" => "\u{a0}",
        "copy" => "©",
        "reg" => "®",
        "trade" => "™",
        "mdash" => "—",
        "ndash" => "–",
        "hellip" => "…",
        "ldquo" => "“",
        "rdquo" => "”",
        "lsquo" => "‘",
        "rsquo" => "’",
        "laquo" => "«",
        "raquo" => "»",
        "bull" => "•",
        "middot" => "·",
        "deg" => "°",
        "plusmn" => "±",
        "times" => "×",
        "divide" => "÷",
        "frac12" => "½",
        "frac14" => "¼",
        "frac34" => "¾",
        "sup2" => "²",
        "sup3" => "³",
        "para" => "¶",
        "sect" => "§",
        "dagger" => "†",
        "euro" => "€",
        "pound" => "£",
        "yen" => "¥",
        "cent" => "¢",
        "larr" => "←",
        "rarr" => "→",
        "harr" => "↔",
        "ne" => "≠",
        "le" => "≤",
        "ge" => "≥",
        _ => return None,
    })
}

/// Escape the interior of every raw-text element, in the SOURCE.
///
/// Inside `<script>` and `<style>`, `<` and `&` are ordinary characters. No
/// setting makes an XML lexer believe that — quick-xml's `read_text` still
/// tokenises what it reads, so `if (a < b)` opened an element named `b)` and
/// the rest of the script vanished into it, silently. The tokenizer is the
/// wrong tool, so the source is fixed before the tokenizer sees it: escape
/// the interior and `decode_text` puts it back on the way out, unchanged.
///
/// This is what the wxhtmledit reference does too — `ExtractStyleBlocks`
/// pulls the blocks out of the HTML string before parsing rather than
/// teaching the parser about them.
fn escape_raw_text(src: &str) -> String {
    let lower = src.to_ascii_lowercase();
    let mut out = String::with_capacity(src.len());
    let mut cursor = 0usize;
    while cursor < src.len() {
        // The next raw-text start tag, whichever comes first.
        let next = RAW_TEXT_ELEMENTS
            .iter()
            .filter_map(|tag| {
                let open = format!("<{tag}");
                lower[cursor..].find(&open).map(|at| (cursor + at, *tag))
            })
            .min_by_key(|(at, _)| *at);
        let Some((start, tag)) = next else { break };
        // Skip the rest of the start tag; an unterminated one is not a
        // raw-text element, it is the end of the input.
        let Some(open_end) = lower[start..].find('>').map(|at| start + at + 1) else {
            break;
        };
        let close = format!("</{tag}");
        let Some(interior_end) = lower[open_end..].find(&close).map(|at| open_end + at) else {
            // Unterminated: everything after it is raw text.
            out.push_str(&src[cursor..open_end]);
            out.push_str(&escape_markup(&src[open_end..]));
            return out;
        };
        out.push_str(&src[cursor..open_end]);
        out.push_str(&escape_markup(&src[open_end..interior_end]));
        cursor = interior_end;
    }
    out.push_str(&src[cursor..]);
    out
}

/// `<` and `&` as character references, so a lexer reads them as text.
fn escape_markup(raw: &str) -> String {
    raw.replace('&', "&amp;").replace('<', "&lt;")
}

/// Decode a text run's character references.
///
/// XML predefines five entities and treats every other name as a fatal
/// error; HTML predefines thousands. `unescape_with` keeps quick-xml's
/// numeric-reference handling and answers named ones from the table above.
fn decode_text(raw: &quick_xml::events::BytesText<'_>, grammar: Grammar) -> String {
    match grammar {
        Grammar::Xml => raw.unescape().map(|c| c.into_owned()).unwrap_or_default(),
        Grammar::Html => raw
            .unescape_with(html_entity)
            .map(|c| c.into_owned())
            // An unresolvable reference must not eat the run it sits in.
            .unwrap_or_else(|_| String::from_utf8_lossy(raw.as_ref()).into_owned()),
    }
}

/// A start tag's element name, folded for HTML.
///
/// HTML tag names are ASCII case-insensitive and `<DIV>` is a `div`. Folding
/// here rather than at each lookup is what keeps everything downstream —
/// `VOID_ELEMENTS`, `implies_end_of`, and in the toolkit `control_kind` and
/// `ua::declarations_for` — able to compare against lowercase literals.
fn element_name(e: &quick_xml::events::BytesStart<'_>, grammar: Grammar) -> String {
    let raw = String::from_utf8_lossy(e.name().as_ref()).into_owned();
    match grammar {
        Grammar::Xml => raw,
        Grammar::Html => raw.to_ascii_lowercase(),
    }
}

fn end_tag_name(e: &quick_xml::events::BytesEnd<'_>) -> String {
    String::from_utf8_lossy(e.name().as_ref())
        .into_owned()
        .to_ascii_lowercase()
}

fn parse_xml(xml: &str) -> Result<Value, String> {
    parse_markup(xml, Grammar::Xml)
}

// ── The parse SINK ────────────────────────────────────────────────────────
//
// **One grammar driver, two trees.**
//
// [`drive`] below is the whole of what this file knows about markup: void
// elements, implied end tags, raw-text interiors, and the tolerant recovery a
// browser performs on input that does not nest. Which TREE that knowledge
// builds is a separate question, and it has two answers:
//
//   * [`ValueSink`] — the property-bag DOM this module has always returned.
//     XML lives here, and so do its consumers: PHP's `DOMDocument`, .NET's
//     XLinq and `xml2js` read `nodeType` / `childNodes` / `attributes` off the
//     object itself, and namespaces, `Attr` nodes, CDATA sections and
//     processing instructions are theirs alone.
//
//   * [`DocumentSink`] — a real document, built through the same `web:dom`
//     operations any guest would use. An HTML page parsed this way IS the kind
//     of thing `window.document` is: it cascades, it lays out, it paints, and
//     every `web:dom` call answers about it. Before this, `parseFromString`
//     returned a tree that could be read and could never be shown.
//
// Splitting the sink rather than the parser is what keeps the HTML algorithm
// from existing twice — which is the shape the two-DOM problem took everywhere
// else it appeared.

/// Where a parse puts what it reads.
trait TreeSink {
    /// Attach an element built from its start tag.
    ///
    /// `open` makes it the parent of what follows; a void or self-closing
    /// element is complete at its start tag and never opens.
    fn start_element(
        &mut self,
        e: &quick_xml::events::BytesStart<'_>,
        grammar: Grammar,
        open: bool,
    );

    /// Close elements until exactly `depth` are open.
    ///
    /// A depth rather than a count of pops, because HTML's end tags close the
    /// nearest MATCHING ancestor: the driver computes which one that is and
    /// says where to stop.
    fn close_to(&mut self, depth: usize);

    fn text(&mut self, data: &str);
    fn cdata(&mut self, data: &str);
    fn comment(&mut self, data: &str);
    fn processing_instruction(&mut self, target: &str, data: &str);

    /// What the SINK could not represent, merged into the driver's own list.
    /// A sink that can hold everything it is handed reports nothing.
    fn notes(&mut self) -> Vec<String> {
        Vec::new()
    }
}

/// Read `source` in `grammar`, building into `sink`.
///
/// Answers what had to be REPAIRED. Once a parser is tolerant, malformed input
/// stops announcing itself with a `<parsererror>` document and starts producing
/// a plausible-but-different tree; this list is the only thing left that says
/// so.
fn drive(source: &str, grammar: Grammar, sink: &mut dyn TreeSink) -> Result<Vec<String>, String> {
    // Raw-text interiors are neutralised before the lexer runs — see
    // `escape_raw_text`. XML has no raw-text elements, so it is untouched.
    let owned;
    let xml = match grammar {
        Grammar::Xml => source,
        Grammar::Html => {
            owned = escape_raw_text(source);
            owned.as_str()
        }
    };
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(false);
    reader.config_mut().expand_empty_elements = false;
    // An end tag that does not match the open element is a fatal error in
    // XML and everyday practice in HTML. Turning the check off is what lets
    // the recovery below run at all — with it on, quick-xml ends the parse
    // before the stack can be unwound to the right element.
    if grammar == Grammar::Html {
        reader.config_mut().check_end_names = false;
        // Separate check, separate flag: `check_end_names` tolerates an end
        // tag that does not match the OPEN element, this one tolerates an
        // end tag matching nothing at all. `</div>` with no `<div>` is a
        // fatal XML error and something a browser simply ignores.
        reader.config_mut().allow_unmatched_ends = true;
    }

    // The open elements, by name. This is the driver's whole state: the sink
    // keeps its own handles in step through `close_to`, and never has to agree
    // with this one on anything but a depth.
    let mut open_names: Vec<String> = Vec::new();
    let mut recoveries: Vec<String> = Vec::new();

    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let name = element_name(&e, grammar);
                // An omitted end tag: `<li>a<li>b` is two siblings, not two
                // nested items. Close every open element the new one ends
                // before the new one is attached, or it lands inside its
                // predecessor.
                if grammar == Grammar::Html {
                    while let Some(open) = open_names.last() {
                        if !implies_end_of(open, &name) {
                            break;
                        }
                        recoveries.push(format!("implied </{open}> before <{name}>"));
                        open_names.pop();
                        sink.close_to(open_names.len());
                    }
                }
                // A void element is complete at its start tag. Opening it
                // would swallow every following sibling as its child.
                let void = grammar == Grammar::Html && VOID_ELEMENTS.contains(&name.as_str());
                sink.start_element(&e, grammar, !void);
                if !void {
                    open_names.push(name);
                }
            }
            Ok(Event::Empty(e)) => {
                // Empty/self-closing — never opened, so siblings continue at
                // the current parent.
                sink.start_element(&e, grammar, false);
            }
            Ok(Event::End(e)) => {
                if grammar == Grammar::Html {
                    let name = end_tag_name(&e);
                    // Close the NEAREST matching ancestor, not whatever is on
                    // top: `<b><i>x</b>` ends the bold, and the italic goes
                    // with it. A browser would reopen the italic afterwards —
                    // that is the adoption agency algorithm, which is out of
                    // scope, so the divergence is recorded rather than hidden.
                    match open_names.iter().rposition(|open| *open == name) {
                        Some(index) => {
                            if index + 1 < open_names.len() {
                                let unclosed = open_names[index + 1..].join(", ");
                                recoveries.push(format!("</{name}> closed still-open {unclosed}"));
                            }
                            open_names.truncate(index);
                            sink.close_to(index);
                        }
                        // An end tag with nothing open to match it is
                        // ignored, which is what a browser does.
                        None => recoveries.push(format!("stray </{name}> ignored")),
                    }
                } else if !open_names.is_empty() {
                    open_names.pop();
                    sink.close_to(open_names.len());
                }
            }
            Ok(Event::Text(t)) => {
                let text = decode_text(&t, grammar);
                if !text.is_empty() {
                    sink.text(&text);
                }
            }
            Ok(Event::CData(c)) => {
                let bytes = c.into_inner().into_owned();
                sink.cdata(&String::from_utf8_lossy(&bytes));
            }
            Ok(Event::Comment(c)) => {
                let text = c.unescape().map(|c| c.into_owned()).unwrap_or_default();
                sink.comment(&text);
            }
            Ok(Event::PI(pi)) => {
                let raw = String::from_utf8_lossy(pi.as_ref()).into_owned();
                // Spec: target = first whitespace-delimited token; data
                // = the rest. `<?xml-stylesheet href="x"?>` → target
                // "xml-stylesheet", data `href="x"`.
                let mut parts = raw.splitn(2, char::is_whitespace);
                let target = parts.next().unwrap_or("").to_string();
                let data = parts.next().unwrap_or("").trim().to_string();
                sink.processing_instruction(&target, &data);
            }
            Ok(Event::Decl(_)) | Ok(Event::DocType(_)) => {
                // XML declaration / DOCTYPE — recorded as DocumentType
                // nodes in Phase 4. For Phase 1 they're skipped (no
                // visible difference for read-side tests).
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(format!("XML parse error: {}", e)),
        }
        buf.clear();
    }

    recoveries.extend(sink.notes());
    Ok(recoveries)
}

/// The property-bag tree — [`make_node`] and friends, driven by [`drive`].
struct ValueSink {
    document: Value,
    /// The document, then every open element. `close_to(n)` truncates to
    /// `n + 1`, because the document is not one of the open elements.
    stack: Vec<Arc<Mutex<Object>>>,
}

impl ValueSink {
    fn new() -> ValueSink {
        let document = make_node(DOCUMENT_NODE, "#document", None);
        let root = match &document {
            Value::Object(o) => o.clone(),
            _ => unreachable!("make_node always answers an object"),
        };
        ValueSink {
            document,
            stack: vec![root],
        }
    }
}

impl TreeSink for ValueSink {
    fn start_element(
        &mut self,
        e: &quick_xml::events::BytesStart<'_>,
        grammar: Grammar,
        open: bool,
    ) {
        let elem = make_element_from_start(e, grammar);
        push_child(&self.stack, elem.clone());
        if open {
            if let Value::Object(o) = elem {
                self.stack.push(o);
            }
        }
    }

    fn close_to(&mut self, depth: usize) {
        self.stack.truncate(depth + 1);
    }

    fn text(&mut self, data: &str) {
        push_child(&self.stack, make_node(TEXT_NODE, "#text", Some(data)));
    }

    fn cdata(&mut self, data: &str) {
        push_child(
            &self.stack,
            make_node(CDATA_SECTION_NODE, "#cdata-section", Some(data)),
        );
    }

    fn comment(&mut self, data: &str) {
        push_child(&self.stack, make_node(COMMENT_NODE, "#comment", Some(data)));
    }

    fn processing_instruction(&mut self, target: &str, data: &str) {
        push_child(&self.stack, make_pi_node(target, data));
    }
}

/// A real document, built through the same `web:dom` operations a guest uses.
///
/// Every line below is `createElement` / `createTextNode` / `createComment` /
/// `setAttribute` / `appendChild` / `textContent` — six standard operations and
/// no seventh. That is deliberate and it is the browser-swap test: the tree
/// this builds is reachable by an engine that has never heard of this parser,
/// because the parser only ever asked it to do things the DOM defines.
struct DocumentSink {
    document: crate::engine::DocumentId,
    /// The lowest depth `close_to` may return to.
    ///
    /// `0` for a whole document — the root is always reachable. For a FRAGMENT
    /// it is `1`: the element being filled sits at index 0 and is not the
    /// fragment's to close.
    floor: usize,
    /// Every open element, innermost last. The document root is not one of
    /// them, so `close_to(0)` returns to it.
    open: Vec<crate::engine::NodeId>,
    /// The open elements' tags, in step with `open`. The sink created them, so
    /// it already knows — reading them back would be asking the tree for a
    /// fact this side just supplied.
    tags: Vec<String>,
    notes: Vec<String>,
}

impl DocumentSink {
    fn new(document: crate::engine::DocumentId) -> DocumentSink {
        DocumentSink {
            document,
            floor: 0,
            open: Vec::new(),
            tags: Vec::new(),
            notes: Vec::new(),
        }
    }

    /// A sink that builds INTO an existing element rather than a new document.
    ///
    /// The whole of what `innerHTML` needs beyond what already existed: the
    /// parser, the grammar and the tree-building were all here, and every entry
    /// point made a fresh `Document`. Seeding the open-element stack with the
    /// target is what points the same machinery at a subtree.
    fn fragment(document: crate::engine::DocumentId, parent: crate::engine::NodeId) -> DocumentSink {
        DocumentSink {
            document,
            floor: 1,
            open: vec![parent],
            tags: vec![String::new()],
            notes: Vec::new(),
        }
    }

    /// The element what-comes-next belongs to.
    fn parent(&self) -> crate::engine::NodeId {
        *self.open.last().unwrap_or(&crate::engine::DOCUMENT)
    }

    fn create(&self, op: DomOp) -> Option<crate::engine::NodeId> {
        match crate::engine::apply(self.document, op) {
            DomValue::Node(node) => Some(node),
            _ => None,
        }
    }

    fn attach(&self, child: crate::engine::NodeId) -> bool {
        matches!(
            crate::engine::apply(
                self.document,
                DomOp::AppendChild {
                    parent: self.parent(),
                    child,
                },
            ),
            DomValue::Bool(true)
        )
    }
}

impl TreeSink for DocumentSink {
    fn start_element(
        &mut self,
        e: &quick_xml::events::BytesStart<'_>,
        grammar: Grammar,
        open: bool,
    ) {
        let name = element_name(e, grammar);
        let attributes: Vec<(String, String)> = e
            .attributes()
            .with_checks(false)
            .flatten()
            .map(|attr| {
                let key = String::from_utf8_lossy(attr.key.as_ref()).to_ascii_lowercase();
                let value = attr
                    .unescape_value_with(html_entity)
                    .map(|c| c.into_owned())
                    .unwrap_or_default();
                (key, value)
            })
            .collect();
        // The attribute that, WITH the tag, decides which control this is —
        // `createElement`'s second argument. Which attribute it is depends on
        // the element: `<input type=checkbox>` and `<select size=6>` are the
        // same fact spelled two ways. The serialiser asks the same question
        // the same way round.
        let disambiguator = match name.as_str() {
            "select" | "datalist" => "size",
            _ => "type",
        };
        let input_type = attributes
            .iter()
            .find(|(key, _)| key == disambiguator)
            .map(|(_, value)| value.clone())
            .unwrap_or_default();

        let Some(node) = self.create(DomOp::CreateElement {
            tag: name.clone(),
            input_type,
        }) else {
            return;
        };
        for (key, value) in &attributes {
            // Already carried into the element as its kind. Setting it again
            // would be the same fact in two places, which is the whole thing
            // this exercise is removing.
            if key == disambiguator {
                continue;
            }
            crate::engine::apply(
                self.document,
                DomOp::SetAttribute(node, key.clone(), value.clone()),
            );
        }
        if !self.attach(node) {
            self.notes
                .push(format!("<{name}> could not join its parent"));
        }
        if open {
            self.open.push(node);
            self.tags.push(name);
        }
    }

    fn close_to(&mut self, depth: usize) {
        // **The driver's depth is relative to ITS root; ours is offset by the
        // seeded element.** A document sink's `open` holds only elements, so
        // driver-depth N is index N. A fragment's holds the element being
        // filled at index 0, so the same N is index N + 1 — and truncating to
        // N closed the wrapper instead of the child inside it.
        //
        // Measured: `<div id=w><div class=h></div><div class=s></div></div>`
        // put `.h` inside `#w` and `.s` beside it, so a one-element-deep
        // fragment silently unwrapped itself after the first child.
        let depth = depth.saturating_add(self.floor);
        self.open.truncate(depth);
        self.tags.truncate(depth);
    }

    fn text(&mut self, data: &str) {
        let parent = self.parent();
        let tag = self.tags.last().cloned().unwrap_or_default();
        // `<title>` names the DOCUMENT, not only the element (HTML §4.2.2).
        if tag == "title" {
            crate::engine::apply(self.document, DomOp::SetTitle(data.to_string()));
        }
        // Inside a raw-text element the run is not content, it is the
        // element's DATA: a `<style>`'s text IS the author stylesheet and a
        // `<script>`'s is its source. `textContent` is the write that says so,
        // and on a `<style>` it is what makes the rules take effect — which is
        // how a parsed page gets a cascade at all.
        if RAW_TEXT_ELEMENTS.contains(&tag.as_str()) {
            crate::engine::apply(
                self.document,
                DomOp::SetTextContent(parent, data.to_string()),
            );
            return;
        }
        if let Some(node) = self.create(DomOp::CreateTextNode(data.to_string())) {
            if self.attach(node) {
                return;
            }
        }
        // Whitespace between two elements is not content and goes nowhere.
        if data.trim().is_empty() {
            return;
        }
        // Text at the root has no box to belong to: the document node is the
        // body, and a body cannot hold a line of its own yet. Said out loud
        // rather than dropped, because it IS a divergence from the markup.
        if parent == crate::engine::DOCUMENT {
            self.notes
                .push(format!("{} characters of root text dropped", data.len()));
            return;
        }
        // A box that refuses a text NODE is a leaf — `<option>`, `<td>`,
        // `<label>`, `<span>`, `<title>` — and for a leaf its text and its
        // `textContent` are the same fact, so this is the same content taking
        // the only shape the element has for it.
        crate::engine::apply(
            self.document,
            DomOp::SetTextContent(parent, data.to_string()),
        );
    }

    fn comment(&mut self, data: &str) {
        if let Some(node) = self.create(DomOp::CreateComment(data.to_string())) {
            self.attach(node);
        }
    }

    /// HTML has no CDATA section outside foreign content: `<![CDATA[x]]>` is a
    /// **bogus comment** whose data is `[CDATA[x]]` (HTML §13.2.5.42). Reusing
    /// the comment node is not an approximation, it is the rule.
    fn cdata(&mut self, data: &str) {
        self.comment(&format!("[CDATA[{data}]]"));
    }

    /// And `<?target data?>` is a bogus comment too, whose data is everything
    /// after the `<?` — HTML has no processing instructions.
    fn processing_instruction(&mut self, target: &str, data: &str) {
        let text = if data.is_empty() {
            format!("?{target}")
        } else {
            format!("?{target} {data}")
        };
        self.comment(&text);
    }

    fn notes(&mut self) -> Vec<String> {
        std::mem::take(&mut self.notes)
    }
}

/// `DOMParser.parseFromString(source, "text/html")`.
///
/// Answers a handle to a **real** document — the same kind of thing
/// `window.document` is, in the same engine, reachable by every `web:dom`
/// operation. A browser returns an `HTMLDocument` here and so does this.
///
/// HTML has no parse errors, so there is no failure branch: what could not be
/// represented comes back as `__parseRecoveries` on the handle.
fn parse_html_document(source: &str) -> Value {
    let document = crate::engine::new_document("");
    let mut sink = DocumentSink::new(document);
    let recoveries = drive(source, Grammar::Html, &mut sink).unwrap_or_default();
    let handle = crate::html::document_handle(document);
    // `__parseRecoveries` is OUR diagnostic, not a DOM member — no browser has
    // one — so it lives on the handle rather than being written into the tree
    // as an attribute no engine would recognise. It is parse-time-only and
    // never changes again, which is why a property here is a record and not
    // the second copy of a live fact.
    if let Value::Object(o) = &handle {
        let entries = recoveries.iter().map(|r| s(r)).collect::<Vec<_>>();
        o.lock()
            .unwrap()
            .properties
            .insert("__parseRecoveries".into(), make_array(entries));
    }
    handle
}

/// **`Element.innerHTML = …`** — parse `source` and make it the element's
/// content, replacing whatever was there (DOM Parsing §2.3).
///
/// Replacement is the spec's own wording and it is not a detail: setting
/// `innerHTML` is defined as discarding the children and inserting the parsed
/// fragment, so appending would leave a page that grows every time it is
/// redrawn — which is exactly what a framework does on each render.
pub fn set_inner_html(document: crate::engine::DocumentId, node: crate::engine::NodeId, source: &str) {
    // Remove first, so a parse that yields nothing still empties the box —
    // `el.innerHTML = ""` is how a page clears itself.
    for child in match crate::engine::apply(document, DomOp::ChildNodes(node)) {
        DomValue::Nodes(children) => children,
        _ => Vec::new(),
    } {
        crate::engine::apply(
            document,
            DomOp::RemoveChild {
                parent: node,
                child,
            },
        );
    }
    let mut sink = DocumentSink::fragment(document, node);
    let _ = drive(source, Grammar::Html, &mut sink);
}

/// Parse `source` into a detached `DocumentFragment` and hand back its id.
///
/// The fragment is what makes the two writes below simple: inserting one
/// splices its children in at a single point, so neither has to walk a list of
/// parsed roots and place them one at a time — and neither can get their order
/// wrong on the way.
fn parse_into_fragment(document: crate::engine::DocumentId, source: &str) -> crate::engine::NodeId {
    let fragment = match crate::engine::apply(document, DomOp::CreateDocumentFragment) {
        DomValue::Node(id) => id,
        _ => return 0,
    };
    let mut sink = DocumentSink::fragment(document, fragment);
    let _ = drive(source, Grammar::Html, &mut sink);
    fragment
}

/// The children of `parent`, or an empty list.
fn children_of(
    document: crate::engine::DocumentId,
    parent: crate::engine::NodeId,
) -> Vec<crate::engine::NodeId> {
    match crate::engine::apply(document, DomOp::ChildNodes(parent)) {
        DomValue::Nodes(children) => children,
        _ => Vec::new(),
    }
}

/// Place `child` under `parent`, before `reference` or at the end.
///
/// **`InsertBefore` cannot express "at the end".** Its `reference` is a bare
/// `NodeId`, and the two engines read a zero one differently — one takes it
/// for the document node, the other for a mistake and does nothing. Neither
/// appends. So the end case has to be `AppendChild`, said explicitly.
fn place(
    document: crate::engine::DocumentId,
    parent: crate::engine::NodeId,
    child: crate::engine::NodeId,
    reference: Option<crate::engine::NodeId>,
) {
    match reference {
        Some(reference) => {
            crate::engine::apply(
                document,
                DomOp::InsertBefore {
                    parent,
                    child,
                    reference,
                },
            );
        }
        None => {
            crate::engine::apply(document, DomOp::AppendChild { parent, child });
        }
    }
}

/// `element.outerHTML = …` — replace the element ITSELF, not its contents.
pub fn set_outer_html(
    document: crate::engine::DocumentId,
    node: crate::engine::NodeId,
    source: &str,
) {
    let parent = match crate::engine::apply(document, DomOp::ParentNode(node)) {
        DomValue::Node(id) => id,
        // A detached element has nowhere to be replaced. The IDL throws here;
        // there is no exception to raise at this seam, and doing nothing is
        // the one outcome that cannot corrupt the tree.
        _ => return,
    };
    let fragment = parse_into_fragment(document, source);
    if fragment != 0 {
        place(document, parent, fragment, Some(node));
    }
    crate::engine::apply(document, DomOp::RemoveChild { parent, child: node });
}

/// `document.importNode(externalNode, deep)`.
///
/// Goes through MARKUP — the source node is serialised in its own document and
/// re-parsed in this one. That is not a shortcut around a real copy: the two
/// documents own separate node tables and separate widgets, so nothing on the
/// source side can be moved or referenced, only described and rebuilt. It is
/// also why the copy carries no event listeners, which is what the spec says a
/// clone does anyway (DOM §4.5, "clone a node" copies no listeners).
///
/// What it does NOT carry is state that lives on a control rather than in an
/// attribute — a typed-into `input.value` with no `value=` attribute is the
/// case. Returns 0 for a node that cannot be described.
pub fn import_node(
    document: crate::engine::DocumentId,
    source: crate::engine::DocumentId,
    node: crate::engine::NodeId,
    deep: bool,
) -> crate::engine::NodeId {
    let markup = match crate::engine::apply(source, DomOp::OuterHtml(node)) {
        DomValue::Text(markup) if !markup.is_empty() => markup,
        _ => return 0,
    };
    let fragment = parse_into_fragment(document, &markup);
    if fragment == 0 {
        return 0;
    }
    // The first ELEMENT, not the first node. Markup that begins with
    // whitespace parses to a leading text node, and returning that would hand
    // back a text node where the caller asked for the element it imported.
    let imported = match children_of(document, fragment)
        .into_iter()
        .find(|child| matches!(crate::engine::apply(document, DomOp::NodeType(*child)),
                               DomValue::Number(kind) if kind == 1.0))
    {
        Some(element) => element,
        None => return 0,
    };
    if !deep {
        // A shallow import is the node and nothing under it. Markup always
        // brings the subtree, so the subtree comes back off.
        for child in children_of(document, imported) {
            crate::engine::apply(
                document,
                DomOp::RemoveChild {
                    parent: imported,
                    child,
                },
            );
        }
    }
    // Detached, as the IDL requires — `importNode` does not insert.
    crate::engine::apply(
        document,
        DomOp::RemoveChild {
            parent: fragment,
            child: imported,
        },
    );
    imported
}

/// `element.insertAdjacentHTML(position, text)`.
pub fn insert_adjacent_html(
    document: crate::engine::DocumentId,
    node: crate::engine::NodeId,
    position: &str,
    source: &str,
) {
    // Where the fragment goes, as a (parent, before) pair. `beforebegin` and
    // `afterend` need the element's PARENT, and a detached element has none —
    // the IDL throws for those two and does not for the other two.
    let (parent, before) = match position.to_ascii_lowercase().as_str() {
        "beforebegin" | "afterend" => {
            let parent = match crate::engine::apply(document, DomOp::ParentNode(node)) {
                DomValue::Node(id) => id,
                _ => return,
            };
            if position.eq_ignore_ascii_case("beforebegin") {
                (parent, Some(node))
            } else {
                // The sibling AFTER this one, found through the parent's child
                // list because the seam has no `nextSibling` op — and `None`
                // when there is none, which appends.
                let siblings = children_of(document, parent);
                let next = siblings
                    .iter()
                    .position(|c| *c == node)
                    .and_then(|at| siblings.get(at + 1))
                    .copied();
                (parent, next)
            }
        }
        "afterbegin" => (node, children_of(document, node).first().copied()),
        "beforeend" => (node, None),
        _ => return,
    };

    let fragment = parse_into_fragment(document, source);
    if fragment == 0 {
        return;
    }
    place(document, parent, fragment, before);
}

fn parse_markup(xml: &str, grammar: Grammar) -> Result<Value, String> {
    let mut sink = ValueSink::new();
    let recoveries = drive(xml, grammar, &mut sink)?;
    let document_obj = sink.document;

    // Post-walk: set parentNode + ownerDocument back-refs, derive
    // textContent / firstChild / lastChild / nextSibling /
    // previousSibling. Element-only children list + first/last
    // ElementChild are also populated here.
    let doc_arc = match &document_obj {
        Value::Object(o) => o.clone(),
        _ => unreachable!(),
    };
    finalize_node_tree(&doc_arc, None, Some(&doc_arc));
    set_document_element(&doc_arc);

    // What the parser repaired, on the document, always — an empty array
    // when nothing was repaired, so "no recoveries" is an answer rather
    // than a missing property.
    if grammar == Grammar::Html {
        let entries = recoveries.iter().map(|r| s(r)).collect::<Vec<_>>();
        doc_arc
            .lock()
            .unwrap()
            .properties
            .insert("__parseRecoveries".into(), make_array(entries));
    }

    Ok(document_obj)
}

fn parse_error_document(xml: &str) -> Value {
    // WHATWG: parseerror documents are still Documents whose root
    // element is `<parsererror>`. We mirror that minimally.
    let doc = make_node(DOCUMENT_NODE, "#document", None);
    let err_elem = {
        let mut o = Object::new();
        o.properties.insert("__type".into(), s("Element"));
        o.properties
            .insert("nodeType".into(), Value::I32(ELEMENT_NODE));
        o.properties.insert("nodeName".into(), s("parsererror"));
        o.properties.insert("tagName".into(), s("parsererror"));
        o.properties.insert("localName".into(), s("parsererror"));
        o.properties.insert("prefix".into(), Value::Null);
        o.properties.insert("namespaceURI".into(), Value::Null);
        o.properties.insert("nodeValue".into(), Value::Null);
        o.properties
            .insert("attributes".into(), make_empty_object());
        o.properties.insert("childNodes".into(), make_array(vec![]));
        o.properties.insert("children".into(), make_array(vec![]));
        o.properties.insert("textContent".into(), s(xml));
        o.properties.insert("id".into(), s(""));
        o.properties.insert("className".into(), s(""));
        Value::Object(vybe_runtime::heap::alloc(o))
    };
    if let Value::Object(d) = &doc {
        if let Value::Object(arr) = d
            .lock()
            .unwrap()
            .properties
            .get("childNodes")
            .cloned()
            .unwrap_or(Value::Null)
        {
            if let ObjectKind::Array(ref mut items) = arr.lock().unwrap().kind {
                items.push(err_elem.clone());
            }
        }
        if let Value::Object(d_obj) = &doc {
            d_obj
                .lock()
                .unwrap()
                .properties
                .insert("documentElement".into(), err_elem);
        }
    }
    if let Value::Object(d) = &doc {
        finalize_node_tree(d, None, Some(d));
    }
    doc
}

// ── Element construction ──────────────────────────────────────────

fn make_element_from_start(e: &quick_xml::events::BytesStart, grammar: Grammar) -> Value {
    let raw_name = element_name(e, grammar);
    // Split optional prefix per Namespaces in XML 1.0 §3.
    let (prefix, local_name) = match raw_name.split_once(':') {
        Some((p, ln)) => (Some(p.to_string()), ln.to_string()),
        None => (None, raw_name.clone()),
    };
    let mut o = Object::new();
    o.type_id = type_id_for(ELEMENT_NODE);
    o.properties.insert("__type".into(), s("Element"));
    o.properties
        .insert("nodeType".into(), Value::I32(ELEMENT_NODE));
    o.properties.insert("nodeName".into(), s(&raw_name));
    o.properties.insert("tagName".into(), s(&raw_name));
    o.properties.insert("localName".into(), s(&local_name));
    o.properties.insert(
        "prefix".into(),
        prefix.map(|p| s(&p)).unwrap_or(Value::Null),
    );
    o.properties.insert("namespaceURI".into(), Value::Null);
    o.properties.insert("nodeValue".into(), Value::Null);

    // Attributes: spec-shaped NamedNodeMap (we use a property bag —
    // close enough for STRUCT_GET-style access; full NamedNodeMap
    // resource arrives in Phase 4).
    let mut attrs = Object::new();
    attrs.properties.insert("__type".into(), s("NamedNodeMap"));
    let mut id_val = String::new();
    let mut class_val = String::new();
    for attr in e.attributes().with_checks(false).flatten() {
        let key = {
            let raw = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
            // HTML attribute names fold too, so `<div CLASS=x>` sets `class`
            // and `getAttribute("class")` finds it.
            match grammar {
                Grammar::Xml => raw,
                Grammar::Html => raw.to_ascii_lowercase(),
            }
        };
        let val = match grammar {
            Grammar::Xml => attr.unescape_value().map(|c| c.into_owned()),
            Grammar::Html => attr.unescape_value_with(html_entity).map(|c| c.into_owned()),
        }
        .unwrap_or_default();
        if key == "id" {
            id_val = val.clone();
        } else if key == "class" {
            class_val = val.clone();
        }
        attrs.properties.insert(key, s(&val));
    }
    o.properties.insert(
        "attributes".into(),
        Value::Object(vybe_runtime::heap::alloc(attrs)),
    );
    o.properties.insert("id".into(), s(&id_val));
    o.properties.insert("className".into(), s(&class_val));

    // childNodes / children populated later as parse progresses.
    o.properties.insert("childNodes".into(), make_array(vec![]));
    o.properties.insert("children".into(), make_array(vec![]));
    Value::Object(vybe_runtime::heap::alloc(o))
}

fn make_node(node_type: i32, node_name: &str, value: Option<&str>) -> Value {
    let mut o = Object::new();
    o.type_id = type_id_for(node_type);
    let typename = match node_type {
        DOCUMENT_NODE => "Document",
        ELEMENT_NODE => "Element",
        TEXT_NODE => "Text",
        CDATA_SECTION_NODE => "CDATASection",
        COMMENT_NODE => "Comment",
        PROCESSING_INSTRUCTION_NODE => "ProcessingInstruction",
        _ => "Node",
    };
    o.properties.insert("__type".into(), s(typename));
    o.properties
        .insert("nodeType".into(), Value::I32(node_type));
    o.properties.insert("nodeName".into(), s(node_name));
    o.properties.insert(
        "nodeValue".into(),
        match value {
            Some(v) => s(v),
            None => Value::Null,
        },
    );
    if node_type == DOCUMENT_NODE {
        o.properties.insert("childNodes".into(), make_array(vec![]));
        o.properties.insert("documentElement".into(), Value::Null);
        o.properties.insert("doctype".into(), Value::Null);
    }
    Value::Object(vybe_runtime::heap::alloc(o))
}

fn make_pi_node(target: &str, data: &str) -> Value {
    let mut o = Object::new();
    o.type_id = type_id_for(PROCESSING_INSTRUCTION_NODE);
    o.properties
        .insert("__type".into(), s("ProcessingInstruction"));
    o.properties
        .insert("nodeType".into(), Value::I32(PROCESSING_INSTRUCTION_NODE));
    o.properties.insert("nodeName".into(), s(target));
    o.properties.insert("target".into(), s(target));
    o.properties.insert("data".into(), s(data));
    o.properties.insert("nodeValue".into(), s(data));
    Value::Object(vybe_runtime::heap::alloc(o))
}

fn push_child(stack: &[Arc<Mutex<Object>>], child: Value) {
    let parent = stack.last().expect("dom parser: empty stack");
    let arr = {
        let parent_lock = parent.lock().unwrap();
        match parent_lock.properties.get("childNodes") {
            Some(Value::Object(arr)) => arr.clone(),
            _ => return,
        }
    };
    let mut a = arr.lock().unwrap();
    let len = if let ObjectKind::Array(ref mut items) = a.kind {
        items.push(child);
        items.len()
    } else {
        return;
    };
    // Keep NodeList.length live on the parse path too (DOM §4.2.10).
    a.properties.insert("length".into(), Value::I32(len as i32));
}

// ── Post-walk: parentNode / ownerDocument / textContent / siblings ──

fn finalize_node_tree(
    node: &Arc<Mutex<Object>>,
    parent: Option<&Arc<Mutex<Object>>>,
    document: Option<&Arc<Mutex<Object>>>,
) {
    {
        let mut n = node.lock().unwrap();
        n.properties.insert(
            "parentNode".into(),
            parent
                .map(|p| Value::Object(p.clone()))
                .unwrap_or(Value::Null),
        );
        n.properties.insert(
            "ownerDocument".into(),
            match document {
                Some(d) if !Arc::ptr_eq(d, node) => Value::Object(d.clone()),
                _ => Value::Null,
            },
        );
    }
    let children = {
        let n = node.lock().unwrap();
        n.properties.get("childNodes").cloned()
    };
    let Some(Value::Object(arr)) = children else {
        return;
    };
    let child_values: Vec<Value> = {
        let a = arr.lock().unwrap();
        if let ObjectKind::Array(ref items) = a.kind {
            items.clone()
        } else {
            return;
        }
    };

    // Recurse into each child.
    for child in &child_values {
        if let Value::Object(child_obj) = child {
            finalize_node_tree(child_obj, Some(node), document);
        }
    }

    // firstChild / lastChild / siblings on the children list.
    for (i, child) in child_values.iter().enumerate() {
        if let Value::Object(child_obj) = child {
            let prev = if i > 0 {
                child_values.get(i - 1).cloned()
            } else {
                None
            };
            let next = child_values.get(i + 1).cloned();
            let mut c = child_obj.lock().unwrap();
            c.properties
                .insert("previousSibling".into(), prev.unwrap_or(Value::Null));
            c.properties
                .insert("nextSibling".into(), next.unwrap_or(Value::Null));
        }
    }

    // Element-only children list + firstElementChild / lastElementChild.
    let element_children: Vec<Value> = child_values
        .iter()
        .filter(|v| {
            matches!(v, Value::Object(o)
            if matches!(o.lock().unwrap().properties.get("nodeType"), Some(Value::I32(1))))
        })
        .cloned()
        .collect();
    let first_elem = element_children.first().cloned().unwrap_or(Value::Null);
    let last_elem = element_children.last().cloned().unwrap_or(Value::Null);

    let mut n = node.lock().unwrap();
    n.properties.insert(
        "firstChild".into(),
        child_values.first().cloned().unwrap_or(Value::Null),
    );
    n.properties.insert(
        "lastChild".into(),
        child_values.last().cloned().unwrap_or(Value::Null),
    );

    let is_element_or_doc = matches!(n.properties.get("nodeType"), Some(Value::I32(node_type))
        if *node_type == ELEMENT_NODE || *node_type == DOCUMENT_NODE);
    let is_text_like = matches!(n.properties.get("nodeType"), Some(Value::I32(t))
        if *t == TEXT_NODE || *t == CDATA_SECTION_NODE);

    if is_element_or_doc {
        // Replace children Array contents in place.
        if let Some(Value::Object(children_arr)) = n.properties.get("children") {
            let children_arr = children_arr.clone();
            let len = element_children.len();
            let mut guard = children_arr.lock().unwrap();
            let replaced = if let ObjectKind::Array(ref mut items) = guard.kind {
                *items = element_children;
                true
            } else {
                false
            };
            // Maintain `HTMLCollection.length` exactly as
            // `refresh_node_relationships` maintains `NodeList.length`: this
            // list was built by `make_array(vec![])` and is filled in place,
            // so the cached `length` property stays 0 unless it is written.
            // Readers split on which they consult — `ecma:array.length` reads
            // the backing vector (right), the array-like iteration fallback
            // reads this property (was stale ⇒ spread drained NOTHING).
            if replaced {
                guard.properties.insert("length".into(), Value::I32(len as i32));
            }
        } else {
            n.properties
                .insert("children".into(), make_array(element_children));
        }
        n.properties.insert("firstElementChild".into(), first_elem);
        n.properties.insert("lastElementChild".into(), last_elem);
    } else if is_text_like {
        let nv = n
            .properties
            .get("nodeValue")
            .cloned()
            .unwrap_or(Value::Null);
        n.properties.insert("textContent".into(), nv);
    }
    // Drop the lock BEFORE recursing into `collect_text`, which
    // re-acquires the same node's mutex. std Mutex isn't reentrant
    // — holding the lock here would deadlock the textContent walk.
    drop(n);

    if is_element_or_doc {
        let mut buf = String::new();
        collect_text(node, &mut buf);
        node.lock()
            .unwrap()
            .properties
            .insert("textContent".into(), s(&buf));
    }
}

fn collect_text(node: &Arc<Mutex<Object>>, out: &mut String) {
    let n = node.lock().unwrap();
    match n.properties.get("nodeType") {
        Some(Value::I32(t)) if *t == TEXT_NODE || *t == CDATA_SECTION_NODE => {
            if let Some(Value::String(s)) = n.properties.get("nodeValue") {
                out.push_str(s.as_ref());
            }
        }
        _ => {
            if let Some(Value::Object(arr)) = n.properties.get("childNodes") {
                let arr = arr.clone();
                drop(n);
                if let ObjectKind::Array(ref items) = arr.lock().unwrap().kind {
                    for child in items {
                        if let Value::Object(child_obj) = child {
                            collect_text(child_obj, out);
                        }
                    }
                }
            }
        }
    }
}

fn set_document_element(doc: &Arc<Mutex<Object>>) {
    let element = {
        let d = doc.lock().unwrap();
        d.properties.get("childNodes").cloned()
    };
    let Some(Value::Object(arr)) = element else {
        return;
    };
    let elem_root = {
        let a = arr.lock().unwrap();
        if let ObjectKind::Array(ref items) = a.kind {
            items
                .iter()
                .find(|v| {
                    matches!(v, Value::Object(o)
                if matches!(o.lock().unwrap().properties.get("nodeType"), Some(Value::I32(1))))
                })
                .cloned()
                .unwrap_or(Value::Null)
        } else {
            Value::Null
        }
    };
    doc.lock()
        .unwrap()
        .properties
        .insert("documentElement".into(), elem_root);
}

// ── Tree walkers (computed methods) ───────────────────────────────

fn find_by_id(node: &Arc<Mutex<Object>>, id: &str) -> Option<Value> {
    let n = node.lock().unwrap();
    // DOM §4.5 `getElementById`: match the element whose `id` ATTRIBUTE equals
    // `id`. Check the mirrored `id` property first, then the attribute map (so
    // parsed elements — which populate attributes, not the property — match).
    let matches = matches!(n.properties.get("id"), Some(Value::String(v)) if v.as_ref() == id)
        || match n.properties.get("attributes") {
            Some(Value::Object(attrs)) => matches!(
                attrs.lock().unwrap().properties.get("id"),
                Some(Value::String(v)) if v.as_ref() == id
            ),
            _ => false,
        };
    if matches {
        return Some(Value::Object(node.clone()));
    }
    let children = n.properties.get("childNodes").cloned();
    drop(n);
    if let Some(Value::Object(arr)) = children {
        let items: Vec<Value> = if let ObjectKind::Array(ref a) = arr.lock().unwrap().kind {
            a.clone()
        } else {
            return None;
        };
        for child in items {
            if let Value::Object(child_obj) = child {
                if let Some(found) = find_by_id(&child_obj, id) {
                    return Some(found);
                }
            }
        }
    }
    None
}

fn find_by_tag_name(node: &Arc<Mutex<Object>>, tag: &str, out: &mut Vec<Value>) {
    let n = node.lock().unwrap();
    if matches!(n.properties.get("nodeType"), Some(Value::I32(t)) if *t == ELEMENT_NODE) {
        if let Some(Value::String(name)) = n.properties.get("tagName") {
            if tag == "*" || name.as_ref() == tag {
                out.push(Value::Object(node.clone()));
            }
        }
    }
    let children = n.properties.get("childNodes").cloned();
    drop(n);
    if let Some(Value::Object(arr)) = children {
        let items: Vec<Value> = if let ObjectKind::Array(ref a) = arr.lock().unwrap().kind {
            a.clone()
        } else {
            return;
        };
        for child in items {
            if let Value::Object(child_obj) = child {
                find_by_tag_name(&child_obj, tag, out);
            }
        }
    }
}

fn find_by_class_name(node: &Arc<Mutex<Object>>, target: &str, out: &mut Vec<Value>) {
    let n = node.lock().unwrap();
    if matches!(n.properties.get("nodeType"), Some(Value::I32(t)) if *t == ELEMENT_NODE) {
        if let Some(Value::String(class_attr)) = n.properties.get("className") {
            if class_attr.split_whitespace().any(|c| c == target) {
                out.push(Value::Object(node.clone()));
            }
        }
    }
    let children = n.properties.get("childNodes").cloned();
    drop(n);
    if let Some(Value::Object(arr)) = children {
        let items: Vec<Value> = if let ObjectKind::Array(ref a) = arr.lock().unwrap().kind {
            a.clone()
        } else {
            return;
        };
        for child in items {
            if let Value::Object(child_obj) = child {
                find_by_class_name(&child_obj, target, out);
            }
        }
    }
}

// ── Serialiser (XMLSerializer.serializeToString) ──────────────────

fn serialize_node(node: &Arc<Mutex<Object>>, out: &mut String) {
    let n = node.lock().unwrap();
    let node_type = match n.properties.get("nodeType") {
        Some(Value::I32(t)) => *t,
        _ => return,
    };
    match node_type {
        DOCUMENT_NODE => {
            let children = n.properties.get("childNodes").cloned();
            drop(n);
            if let Some(Value::Object(arr)) = children {
                let items: Vec<Value> = if let ObjectKind::Array(ref a) = arr.lock().unwrap().kind {
                    a.clone()
                } else {
                    return;
                };
                for child in items {
                    if let Value::Object(c) = child {
                        serialize_node(&c, out);
                    }
                }
            }
        }
        ELEMENT_NODE => {
            let tag = match n.properties.get("tagName") {
                Some(Value::String(s)) => s.to_string(),
                _ => "".to_string(),
            };
            out.push('<');
            out.push_str(&tag);
            // Attributes
            if let Some(Value::Object(attrs)) = n.properties.get("attributes") {
                let attrs_lock = attrs.lock().unwrap();
                for (k, v) in &attrs_lock.properties {
                    if k.starts_with("__") {
                        continue;
                    }
                    out.push(' ');
                    out.push_str(k);
                    out.push_str("=\"");
                    if let Value::String(vs) = v {
                        escape_attr_into(vs.as_ref(), out);
                    }
                    out.push('"');
                }
            }
            let children = n.properties.get("childNodes").cloned();
            drop(n);
            let items: Vec<Value> = match children {
                Some(Value::Object(arr)) => {
                    if let ObjectKind::Array(ref a) = arr.lock().unwrap().kind {
                        a.clone()
                    } else {
                        vec![]
                    }
                }
                _ => vec![],
            };
            if items.is_empty() {
                out.push_str("/>");
            } else {
                out.push('>');
                for child in items {
                    if let Value::Object(c) = child {
                        serialize_node(&c, out);
                    }
                }
                out.push_str("</");
                out.push_str(&tag);
                out.push('>');
            }
        }
        TEXT_NODE => {
            if let Some(Value::String(s)) = n.properties.get("nodeValue") {
                escape_text_into(s.as_ref(), out);
            }
        }
        CDATA_SECTION_NODE => {
            out.push_str("<![CDATA[");
            if let Some(Value::String(s)) = n.properties.get("nodeValue") {
                out.push_str(s.as_ref());
            }
            out.push_str("]]>");
        }
        COMMENT_NODE => {
            out.push_str("<!--");
            if let Some(Value::String(s)) = n.properties.get("nodeValue") {
                out.push_str(s.as_ref());
            }
            out.push_str("-->");
        }
        PROCESSING_INSTRUCTION_NODE => {
            out.push_str("<?");
            if let Some(Value::String(s)) = n.properties.get("target") {
                out.push_str(s.as_ref());
            }
            if let Some(Value::String(s)) = n.properties.get("data") {
                if !s.is_empty() {
                    out.push(' ');
                    out.push_str(s.as_ref());
                }
            }
            out.push_str("?>");
        }
        _ => {}
    }
}

fn escape_text_into(s: &str, out: &mut String) {
    for c in s.chars() {
        match c {
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            '&' => out.push_str("&amp;"),
            other => out.push(other),
        }
    }
}

fn escape_attr_into(s: &str, out: &mut String) {
    for c in s.chars() {
        match c {
            '"' => out.push_str("&quot;"),
            '<' => out.push_str("&lt;"),
            '&' => out.push_str("&amp;"),
            other => out.push(other),
        }
    }
}

// ── Tiny helpers ──────────────────────────────────────────────────

fn s(v: &str) -> Value {
    Value::String(Arc::from(v))
}

fn make_array(elems: Vec<Value>) -> Value {
    Value::Object(vybe_runtime::heap::alloc(Object::new_array(elems)))
}

fn empty_array() -> Value {
    make_array(vec![])
}

fn make_empty_object() -> Value {
    let mut o = Object::new();
    o.properties.insert("__type".into(), s("NamedNodeMap"));
    Value::Object(vybe_runtime::heap::alloc(o))
}

// ── Selectors API Level 1 — parser + matcher ──────────────────────
// Implements a subset of the CSS Selectors Level 4 syntax sufficient
// for practical XML/HTML access:
//
//   simple_selector := type | universal | id | class | attribute_selector
//   compound       := simple_selector+
//   combinator     := ' ' | '>' | '+' | '~'
//   complex        := compound (combinator compound)*
//   selector_list  := complex (',' complex)*
//
// Attribute matchers:
//   [attr]           — has attribute
//   [attr="value"]   — exact match
//   [attr*="value"]  — substring match
//   [attr^="value"]  — prefix match
//   [attr$="value"]  — suffix match
//   [attr~="value"]  — whitespace-separated word match
//   [attr|="value"]  — exact or hyphen-prefix match

#[derive(Debug, Clone)]
enum SimplePart {
    Universal,
    Type(String),
    Id(String),
    Class(String),
    Attr {
        name: String,
        op: AttrOp,
        value: Option<String>,
    },
}

#[derive(Debug, Clone)]
enum AttrOp {
    Has,
    Exact,
    Substring,
    Prefix,
    Suffix,
    Word,
    Lang,
}

#[derive(Debug, Clone)]
struct CompoundSelector {
    parts: Vec<SimplePart>,
}

#[derive(Debug, Clone)]
enum Combinator {
    Descendant,
    Child,
    AdjacentSibling,
    GeneralSibling,
}

#[derive(Debug, Clone)]
struct ComplexSelector {
    /// Pairs of (combinator-from-previous, compound). The first entry's
    /// combinator is unused (root of the chain).
    parts: Vec<(Combinator, CompoundSelector)>,
}

fn parse_selector_list(input: &str) -> Option<Vec<ComplexSelector>> {
    let mut out = Vec::new();
    for piece in input.split(',') {
        let trimmed = piece.trim();
        if trimmed.is_empty() {
            return None;
        }
        out.push(parse_complex(trimmed)?);
    }
    if out.is_empty() { None } else { Some(out) }
}

fn parse_complex(input: &str) -> Option<ComplexSelector> {
    // Tokenise into (combinator, compound) pairs.
    let chars: Vec<char> = input.chars().collect();
    let mut i = 0;
    let mut parts: Vec<(Combinator, CompoundSelector)> = Vec::new();
    let mut next_combinator = Combinator::Descendant; // unused for the first compound
    let mut first = true;

    while i < chars.len() {
        // Skip whitespace; remember if we crossed any (descendant combinator).
        let pre_ws_start = i;
        while i < chars.len() && chars[i].is_whitespace() {
            i += 1;
        }
        let had_whitespace = i > pre_ws_start;
        if i >= chars.len() {
            break;
        }
        // Explicit combinator?
        let explicit = match chars[i] {
            '>' => Some(Combinator::Child),
            '+' => Some(Combinator::AdjacentSibling),
            '~' => Some(Combinator::GeneralSibling),
            _ => None,
        };
        if let Some(comb) = explicit {
            next_combinator = comb;
            i += 1;
            // Skip following whitespace.
            while i < chars.len() && chars[i].is_whitespace() {
                i += 1;
            }
        } else if had_whitespace && !first {
            next_combinator = Combinator::Descendant;
        }

        let (compound, consumed) = parse_compound(&chars[i..])?;
        if consumed == 0 {
            return None;
        }
        i += consumed;

        if first {
            parts.push((Combinator::Descendant, compound));
            first = false;
        } else {
            parts.push((next_combinator.clone(), compound));
        }
    }

    if parts.is_empty() {
        None
    } else {
        Some(ComplexSelector { parts })
    }
}

fn parse_compound(chars: &[char]) -> Option<(CompoundSelector, usize)> {
    let mut parts: Vec<SimplePart> = Vec::new();
    let mut i = 0;
    while i < chars.len() {
        let c = chars[i];
        if c.is_whitespace() || c == '>' || c == '+' || c == '~' || c == ',' {
            break;
        }
        if c == '*' {
            parts.push(SimplePart::Universal);
            i += 1;
        } else if c == '#' {
            i += 1;
            let start = i;
            while i < chars.len() && is_ident_char(chars[i]) {
                i += 1;
            }
            if start == i {
                return None;
            }
            parts.push(SimplePart::Id(chars[start..i].iter().collect()));
        } else if c == '.' {
            i += 1;
            let start = i;
            while i < chars.len() && is_ident_char(chars[i]) {
                i += 1;
            }
            if start == i {
                return None;
            }
            parts.push(SimplePart::Class(chars[start..i].iter().collect()));
        } else if c == '[' {
            i += 1;
            let (attr, consumed) = parse_attr(&chars[i..])?;
            i += consumed;
            parts.push(attr);
        } else if is_ident_start(c) {
            let start = i;
            while i < chars.len() && is_ident_char(chars[i]) {
                i += 1;
            }
            parts.push(SimplePart::Type(chars[start..i].iter().collect()));
        } else {
            // Unknown character — bail.
            return None;
        }
    }
    if parts.is_empty() {
        None
    } else {
        Some((CompoundSelector { parts }, i))
    }
}

fn parse_attr(chars: &[char]) -> Option<(SimplePart, usize)> {
    let mut i = 0;
    // attr name
    let start = i;
    while i < chars.len() && is_ident_char(chars[i]) {
        i += 1;
    }
    if start == i {
        return None;
    }
    let name: String = chars[start..i].iter().collect();
    // Skip whitespace
    while i < chars.len() && chars[i].is_whitespace() {
        i += 1;
    }
    if i >= chars.len() {
        return None;
    }
    if chars[i] == ']' {
        return Some((
            SimplePart::Attr {
                name,
                op: AttrOp::Has,
                value: None,
            },
            i + 1,
        ));
    }
    // Operator: =, *=, ^=, $=, ~=, |=
    let op = match chars[i] {
        '=' => {
            i += 1;
            AttrOp::Exact
        }
        '*' if i + 1 < chars.len() && chars[i + 1] == '=' => {
            i += 2;
            AttrOp::Substring
        }
        '^' if i + 1 < chars.len() && chars[i + 1] == '=' => {
            i += 2;
            AttrOp::Prefix
        }
        '$' if i + 1 < chars.len() && chars[i + 1] == '=' => {
            i += 2;
            AttrOp::Suffix
        }
        '~' if i + 1 < chars.len() && chars[i + 1] == '=' => {
            i += 2;
            AttrOp::Word
        }
        '|' if i + 1 < chars.len() && chars[i + 1] == '=' => {
            i += 2;
            AttrOp::Lang
        }
        _ => return None,
    };
    // Skip whitespace
    while i < chars.len() && chars[i].is_whitespace() {
        i += 1;
    }
    if i >= chars.len() {
        return None;
    }
    // Quoted or unquoted value
    let value = if chars[i] == '"' || chars[i] == '\'' {
        let quote = chars[i];
        i += 1;
        let start = i;
        while i < chars.len() && chars[i] != quote {
            i += 1;
        }
        if i >= chars.len() {
            return None;
        }
        let v: String = chars[start..i].iter().collect();
        i += 1; // closing quote
        v
    } else {
        let start = i;
        while i < chars.len() && is_ident_char(chars[i]) {
            i += 1;
        }
        chars[start..i].iter().collect()
    };
    // Skip whitespace + closing ']'.
    while i < chars.len() && chars[i].is_whitespace() {
        i += 1;
    }
    if i >= chars.len() || chars[i] != ']' {
        return None;
    }
    Some((
        SimplePart::Attr {
            name,
            op,
            value: Some(value),
        },
        i + 1,
    ))
}

fn is_ident_start(c: char) -> bool {
    c.is_ascii_alphabetic() || c == '_' || c == '-' || c == '\\' || (c as u32) > 0x7F
}
fn is_ident_char(c: char) -> bool {
    c.is_ascii_alphanumeric() || c == '_' || c == '-' || c == '\\' || (c as u32) > 0x7F
}

// ── Matching ─────────────────────────────────────────────────────

fn selector_matches(elem: &Arc<Mutex<Object>>, selector: &ComplexSelector) -> bool {
    // Match right-to-left: the last compound must match `elem`; each
    // earlier compound matches an ancestor / sibling per its combinator
    // to the right neighbour.
    let parts = &selector.parts;
    if parts.is_empty() {
        return true;
    }
    if !compound_matches(elem, &parts.last().unwrap().1) {
        return false;
    }
    let mut current = elem.clone();
    let mut idx = parts.len() - 1;
    while idx > 0 {
        let (combinator, _) = &parts[idx];
        // Walk per combinator until we find a matching ancestor/sibling
        // OR exhaust the chain.
        let prev_compound = &parts[idx - 1].1;
        match combinator {
            Combinator::Descendant => {
                let mut found = false;
                let mut node = parent_of(&current);
                while let Some(p) = node {
                    if compound_matches(&p, prev_compound) {
                        current = p;
                        found = true;
                        break;
                    }
                    node = parent_of(&current_safe(&p));
                    if let Some(pp) = node.clone() {
                        current = pp.clone();
                    } else {
                        break;
                    }
                    node = parent_of(&current);
                }
                if !found {
                    return false;
                }
            }
            Combinator::Child => {
                let parent = match parent_of(&current) {
                    Some(p) => p,
                    None => return false,
                };
                if !compound_matches(&parent, prev_compound) {
                    return false;
                }
                current = parent;
            }
            Combinator::AdjacentSibling => {
                let prev_sibling = previous_sibling_of(&current);
                let p = match prev_sibling {
                    Some(p) => p,
                    None => return false,
                };
                if !compound_matches(&p, prev_compound) {
                    return false;
                }
                current = p;
            }
            Combinator::GeneralSibling => {
                let mut found = false;
                let mut sibling = previous_sibling_of(&current);
                while let Some(s) = sibling {
                    if compound_matches(&s, prev_compound) {
                        current = s;
                        found = true;
                        break;
                    }
                    sibling = previous_sibling_of(&current_safe(&s));
                }
                if !found {
                    return false;
                }
            }
        }
        idx -= 1;
    }
    true
}

fn compound_matches(elem: &Arc<Mutex<Object>>, compound: &CompoundSelector) -> bool {
    let n = elem.lock().unwrap();
    if !matches!(n.properties.get("nodeType"), Some(Value::I32(t)) if *t == ELEMENT_NODE) {
        return false;
    }
    let tag_name = match n.properties.get("tagName") {
        Some(Value::String(s)) => s.to_string(),
        _ => String::new(),
    };
    let id = match n.properties.get("id") {
        Some(Value::String(s)) => s.to_string(),
        _ => String::new(),
    };
    let class_attr = match n.properties.get("className") {
        Some(Value::String(s)) => s.to_string(),
        _ => String::new(),
    };
    let attrs = n.properties.get("attributes").cloned();
    drop(n);

    for part in &compound.parts {
        match part {
            SimplePart::Universal => {}
            SimplePart::Type(t) => {
                if !tag_name.eq_ignore_ascii_case(t) {
                    return false;
                }
            }
            SimplePart::Id(target) => {
                if id != *target {
                    return false;
                }
            }
            SimplePart::Class(target) => {
                if !class_attr.split_whitespace().any(|c| c == target) {
                    return false;
                }
            }
            SimplePart::Attr { name, op, value } => {
                let v = match &attrs {
                    Some(Value::Object(a)) => {
                        let lock = a.lock().unwrap();
                        match lock.properties.get(name) {
                            Some(Value::String(s)) => Some(s.to_string()),
                            Some(other) => Some(format!("{}", other)),
                            None => None,
                        }
                    }
                    _ => None,
                };
                let actual = match v {
                    Some(v) => v,
                    None => return false,
                };
                let target = value.as_deref().unwrap_or("");
                let ok = match op {
                    AttrOp::Has => true,
                    AttrOp::Exact => actual == target,
                    AttrOp::Substring => actual.contains(target),
                    AttrOp::Prefix => actual.starts_with(target),
                    AttrOp::Suffix => actual.ends_with(target),
                    AttrOp::Word => actual.split_whitespace().any(|w| w == target),
                    AttrOp::Lang => actual == target || actual.starts_with(&format!("{}-", target)),
                };
                if !ok {
                    return false;
                }
            }
        }
    }
    true
}

fn parent_of(node: &Arc<Mutex<Object>>) -> Option<Arc<Mutex<Object>>> {
    let n = node.lock().unwrap();
    match n.properties.get("parentNode") {
        Some(Value::Object(p)) => Some(p.clone()),
        _ => None,
    }
}

fn previous_sibling_of(node: &Arc<Mutex<Object>>) -> Option<Arc<Mutex<Object>>> {
    // Element-only previous sibling — required by CSS adjacent / general
    // sibling semantics. We walk the `previousSibling` chain and stop at
    // the first Element.
    let mut cur = {
        let n = node.lock().unwrap();
        match n.properties.get("previousSibling") {
            Some(Value::Object(p)) => Some(p.clone()),
            _ => None,
        }
    };
    while let Some(c) = cur {
        let nt = {
            let n = c.lock().unwrap();
            n.properties.get("nodeType").cloned()
        };
        if let Some(Value::I32(t)) = nt {
            if t == ELEMENT_NODE {
                return Some(c);
            }
        }
        cur = {
            let n = c.lock().unwrap();
            match n.properties.get("previousSibling") {
                Some(Value::Object(p)) => Some(p.clone()),
                _ => None,
            }
        };
    }
    None
}

fn current_safe(node: &Arc<Mutex<Object>>) -> Arc<Mutex<Object>> {
    node.clone()
}

fn find_first_match(root: &Arc<Mutex<Object>>, selectors: &[ComplexSelector]) -> Option<Value> {
    if selectors.iter().any(|s| selector_matches(root, s)) && {
        let n = root.lock().unwrap();
        matches!(n.properties.get("nodeType"), Some(Value::I32(t)) if *t == ELEMENT_NODE)
    } {
        return Some(Value::Object(root.clone()));
    }
    let children = {
        let n = root.lock().unwrap();
        n.properties.get("childNodes").cloned()
    };
    let Some(Value::Object(arr)) = children else {
        return None;
    };
    let items: Vec<Value> = if let ObjectKind::Array(ref a) = arr.lock().unwrap().kind {
        a.clone()
    } else {
        return None;
    };
    for child in items {
        if let Value::Object(c) = child {
            if let Some(found) = find_first_match(&c, selectors) {
                return Some(found);
            }
        }
    }
    None
}

fn find_all_matches(
    root: &Arc<Mutex<Object>>,
    selectors: &[ComplexSelector],
    out: &mut Vec<Value>,
) {
    let is_element = {
        let n = root.lock().unwrap();
        matches!(n.properties.get("nodeType"), Some(Value::I32(t)) if *t == ELEMENT_NODE)
    };
    if is_element && selectors.iter().any(|s| selector_matches(root, s)) {
        out.push(Value::Object(root.clone()));
    }
    let children = {
        let n = root.lock().unwrap();
        n.properties.get("childNodes").cloned()
    };
    let Some(Value::Object(arr)) = children else {
        return;
    };
    let items: Vec<Value> = if let ObjectKind::Array(ref a) = arr.lock().unwrap().kind {
        a.clone()
    } else {
        return;
    };
    for child in items {
        if let Value::Object(c) = child {
            find_all_matches(&c, selectors, out);
        }
    }
}

// ── Mutation helpers ──────────────────────────────────────────────

fn new_element_node(tag: &str, owner: Option<&Arc<Mutex<Object>>>) -> Value {
    // Split optional prefix per Namespaces in XML 1.0 §3.
    let (prefix, local_name) = match tag.split_once(':') {
        Some((p, ln)) => (Some(p.to_string()), ln.to_string()),
        None => (None, tag.to_string()),
    };
    let mut o = Object::new();
    o.type_id = type_id_for(ELEMENT_NODE);
    o.properties.insert("__type".into(), s("Element"));
    o.properties
        .insert("nodeType".into(), Value::I32(ELEMENT_NODE));
    o.properties.insert("nodeName".into(), s(tag));
    o.properties.insert("tagName".into(), s(tag));
    o.properties.insert("localName".into(), s(&local_name));
    o.properties.insert(
        "prefix".into(),
        prefix.map(|p| s(&p)).unwrap_or(Value::Null),
    );
    o.properties.insert("namespaceURI".into(), Value::Null);
    o.properties.insert("nodeValue".into(), Value::Null);

    let mut attrs = Object::new();
    attrs.type_id = dom_type_ids().map(|d| d.named_node_map).unwrap_or(0);
    attrs.properties.insert("__type".into(), s("NamedNodeMap"));
    o.properties.insert(
        "attributes".into(),
        Value::Object(vybe_runtime::heap::alloc(attrs)),
    );
    o.properties.insert("id".into(), s(""));
    o.properties.insert("className".into(), s(""));
    o.properties.insert("childNodes".into(), make_array(vec![]));
    o.properties.insert("children".into(), make_array(vec![]));
    o.properties.insert("textContent".into(), s(""));
    o.properties.insert("parentNode".into(), Value::Null);
    o.properties.insert("firstChild".into(), Value::Null);
    o.properties.insert("lastChild".into(), Value::Null);
    o.properties.insert("nextSibling".into(), Value::Null);
    o.properties.insert("previousSibling".into(), Value::Null);
    o.properties.insert("firstElementChild".into(), Value::Null);
    o.properties.insert("lastElementChild".into(), Value::Null);
    o.properties.insert(
        "ownerDocument".into(),
        owner
            .map(|d| Value::Object(d.clone()))
            .unwrap_or(Value::Null),
    );
    Value::Object(vybe_runtime::heap::alloc(o))
}

fn detach_from_parent(node: &Arc<Mutex<Object>>) {
    let parent = {
        let n = node.lock().unwrap();
        match n.properties.get("parentNode") {
            Some(Value::Object(p)) => Some(p.clone()),
            _ => None,
        }
    };
    if let Some(p) = parent {
        let _ = remove_child_inner(&p, node);
    }
}

fn append_child_inner(parent: &Arc<Mutex<Object>>, child: &Arc<Mutex<Object>>) {
    // Push to childNodes; rewire siblings; refresh parent's
    // children/firstChild/lastChild/textContent.
    let arr = {
        let p = parent.lock().unwrap();
        p.properties.get("childNodes").cloned()
    };
    let Some(Value::Object(arr)) = arr else {
        return;
    };
    {
        let mut a = arr.lock().unwrap();
        if let ObjectKind::Array(ref mut items) = a.kind {
            items.push(Value::Object(child.clone()));
        }
    }
    // Update child's parent.
    {
        let mut c = child.lock().unwrap();
        c.properties
            .insert("parentNode".into(), Value::Object(parent.clone()));
    }
    refresh_node_relationships(parent);
}

fn remove_child_inner(parent: &Arc<Mutex<Object>>, child: &Arc<Mutex<Object>>) -> bool {
    let arr = {
        let p = parent.lock().unwrap();
        p.properties.get("childNodes").cloned()
    };
    let Some(Value::Object(arr)) = arr else {
        return false;
    };
    let removed = {
        let mut a = arr.lock().unwrap();
        let mut found = false;
        if let ObjectKind::Array(ref mut items) = a.kind {
            let target = Arc::as_ptr(child);
            items.retain(|v| {
                if let Value::Object(o) = v {
                    if Arc::as_ptr(o) == target {
                        found = true;
                        return false;
                    }
                }
                true
            });
        }
        found
    };
    if removed {
        // Detach child references.
        let mut c = child.lock().unwrap();
        c.properties.insert("parentNode".into(), Value::Null);
        c.properties.insert("nextSibling".into(), Value::Null);
        c.properties.insert("previousSibling".into(), Value::Null);
        drop(c);
        refresh_node_relationships(parent);
    }
    removed
}

fn insert_before_inner(
    parent: &Arc<Mutex<Object>>,
    new_child: &Arc<Mutex<Object>>,
    reference: &Arc<Mutex<Object>>,
) -> bool {
    let arr = {
        let p = parent.lock().unwrap();
        p.properties.get("childNodes").cloned()
    };
    let Some(Value::Object(arr)) = arr else {
        return false;
    };
    let inserted = {
        let mut a = arr.lock().unwrap();
        let mut idx = None;
        if let ObjectKind::Array(ref items) = a.kind {
            let target = Arc::as_ptr(reference);
            for (i, v) in items.iter().enumerate() {
                if let Value::Object(o) = v {
                    if Arc::as_ptr(o) == target {
                        idx = Some(i);
                        break;
                    }
                }
            }
        }
        if let Some(i) = idx {
            if let ObjectKind::Array(ref mut items) = a.kind {
                items.insert(i, Value::Object(new_child.clone()));
            }
            true
        } else {
            false
        }
    };
    if inserted {
        new_child
            .lock()
            .unwrap()
            .properties
            .insert("parentNode".into(), Value::Object(parent.clone()));
        refresh_node_relationships(parent);
    }
    inserted
}

/// Rebuild firstChild / lastChild / nextSibling / previousSibling
/// pointers + the Element-only `children` list + textContent on the
/// parent and its now-current child set. Called after any mutation.
fn refresh_node_relationships(parent: &Arc<Mutex<Object>>) {
    let children = {
        let p = parent.lock().unwrap();
        p.properties.get("childNodes").cloned()
    };
    let Some(Value::Object(arr)) = children else {
        return;
    };
    let items: Vec<Value> = if let ObjectKind::Array(ref a) = arr.lock().unwrap().kind {
        a.clone()
    } else {
        return;
    };

    // Maintain `NodeList.length` (DOM §4.2.10) on the childNodes list object
    // so `node.childNodes.length` reads the live count.
    {
        arr.lock()
            .unwrap()
            .properties
            .insert("length".into(), Value::I32(items.len() as i32));
    }

    for (i, child) in items.iter().enumerate() {
        if let Value::Object(child_obj) = child {
            let prev = if i > 0 {
                items.get(i - 1).cloned()
            } else {
                None
            };
            let next = items.get(i + 1).cloned();
            let mut c = child_obj.lock().unwrap();
            c.properties
                .insert("previousSibling".into(), prev.unwrap_or(Value::Null));
            c.properties
                .insert("nextSibling".into(), next.unwrap_or(Value::Null));
        }
    }

    let element_children: Vec<Value> = items
        .iter()
        .filter(|v| {
            matches!(v, Value::Object(o)
            if matches!(o.lock().unwrap().properties.get("nodeType"), Some(Value::I32(1))))
        })
        .cloned()
        .collect();

    let first_child = items.first().cloned().unwrap_or(Value::Null);
    let last_child = items.last().cloned().unwrap_or(Value::Null);
    let first_elem = element_children.first().cloned().unwrap_or(Value::Null);
    let last_elem = element_children.last().cloned().unwrap_or(Value::Null);

    {
        let mut p = parent.lock().unwrap();
        p.properties.insert("firstChild".into(), first_child);
        p.properties.insert("lastChild".into(), last_child);
        // Element / Document / DocumentFragment all expose Element-only
        // helpers (`children`, `firstElementChild`, `lastElementChild`).
        // 11 = DOCUMENT_FRAGMENT_NODE.
        let is_elem_or_doc = matches!(p.properties.get("nodeType"), Some(Value::I32(t))
            if *t == ELEMENT_NODE || *t == DOCUMENT_NODE || *t == 11);
        if is_elem_or_doc {
            if let Some(Value::Object(children_arr)) = p.properties.get("children") {
                let children_arr = children_arr.clone();
                let len = element_children.len();
                let mut guard = children_arr.lock().unwrap();
                let replaced = if let ObjectKind::Array(ref mut a_items) = guard.kind {
                    *a_items = element_children;
                    true
                } else {
                    false
                };
                // `HTMLCollection.length`, the mirror of the `NodeList.length`
                // write above. Every live mutator (appendChild / removeChild /
                // insertBefore / fragment move) reaches `children` through
                // here, so this one write covers all of them.
                if replaced {
                    guard.properties.insert("length".into(), Value::I32(len as i32));
                }
            } else {
                p.properties
                    .insert("children".into(), make_array(element_children));
            }
            p.properties.insert("firstElementChild".into(), first_elem);
            p.properties.insert("lastElementChild".into(), last_elem);
            // documentElement on Document always tracks first Element child.
            if matches!(p.properties.get("nodeType"), Some(Value::I32(t)) if *t == DOCUMENT_NODE) {
                let elem_root = items
                    .iter()
                    .find(|v| {
                        matches!(v, Value::Object(o)
                    if matches!(o.lock().unwrap().properties.get("nodeType"), Some(Value::I32(1))))
                    })
                    .cloned()
                    .unwrap_or(Value::Null);
                p.properties.insert("documentElement".into(), elem_root);
            }
        }
    }
    // textContent on the parent — needs lock dropped before recursing.
    let mut buf = String::new();
    collect_text(parent, &mut buf);
    parent
        .lock()
        .unwrap()
        .properties
        .insert("textContent".into(), s(&buf));
}

fn clone_node(node: &Arc<Mutex<Object>>, deep: bool) -> Value {
    let n = node.lock().unwrap();
    let mut clone = Object::new();
    clone.type_id = n.type_id;
    // Shallow copy of all properties EXCEPT structural pointers.
    for (key, val) in &n.properties {
        match key.as_str() {
            "parentNode" | "nextSibling" | "previousSibling" | "firstChild" | "lastChild"
            | "firstElementChild" | "lastElementChild" | "ownerDocument" => continue, // reset structural state
            "childNodes" | "children" | "attributes" => continue, // handled below
            _ => {
                clone.properties.insert(key.clone(), val.clone());
            }
        }
    }
    // Reset structural state.
    clone.properties.insert("parentNode".into(), Value::Null);
    clone.properties.insert("nextSibling".into(), Value::Null);
    clone
        .properties
        .insert("previousSibling".into(), Value::Null);
    clone.properties.insert("firstChild".into(), Value::Null);
    clone.properties.insert("lastChild".into(), Value::Null);
    clone
        .properties
        .insert("firstElementChild".into(), Value::Null);
    clone
        .properties
        .insert("lastElementChild".into(), Value::Null);
    clone.properties.insert("ownerDocument".into(), Value::Null);

    // Clone attributes (deep copy of NamedNodeMap).
    if let Some(Value::Object(attrs_src)) = n.properties.get("attributes") {
        let mut attrs_clone = Object::new();
        attrs_clone.type_id = attrs_src.lock().unwrap().type_id;
        for (k, v) in &attrs_src.lock().unwrap().properties {
            attrs_clone.properties.insert(k.clone(), v.clone());
        }
        clone.properties.insert(
            "attributes".into(),
            Value::Object(vybe_runtime::heap::alloc(attrs_clone)),
        );
    }
    clone
        .properties
        .insert("childNodes".into(), make_array(vec![]));
    clone
        .properties
        .insert("children".into(), make_array(vec![]));
    drop(n);

    let clone_arc = vybe_runtime::heap::alloc(clone);
    if deep {
        let kids = {
            let n = node.lock().unwrap();
            n.properties.get("childNodes").cloned()
        };
        if let Some(Value::Object(arr)) = kids {
            let items: Vec<Value> = if let ObjectKind::Array(ref a) = arr.lock().unwrap().kind {
                a.clone()
            } else {
                vec![]
            };
            for child in items {
                if let Value::Object(c) = child {
                    let cloned_child = clone_node(&c, true);
                    if let Value::Object(cc) = &cloned_child {
                        append_child_inner(&clone_arc, cc);
                    }
                }
            }
        }
    }
    Value::Object(clone_arc)
}

// ── Phase 4 helpers ───────────────────────────────────────────────

fn new_document_fragment(owner: Option<&Arc<Mutex<Object>>>) -> Value {
    // DocumentFragment shares enough of Element's shape that we treat
    // it as a generic node with `nodeType=11` and `nodeName="#document-fragment"`.
    // Method dispatch routes through the same `Element` vtable for
    // appendChild/etc. since this class isn't separately registered.
    const DOCUMENT_FRAGMENT_NODE: i32 = 11;
    let mut o = Object::new();
    o.type_id = dom_type_ids().map(|d| d.element).unwrap_or(0);
    o.properties.insert("__type".into(), s("DocumentFragment"));
    o.properties
        .insert("nodeType".into(), Value::I32(DOCUMENT_FRAGMENT_NODE));
    o.properties
        .insert("nodeName".into(), s("#document-fragment"));
    o.properties.insert("nodeValue".into(), Value::Null);
    o.properties.insert("childNodes".into(), make_array(vec![]));
    o.properties.insert("children".into(), make_array(vec![]));
    o.properties.insert("textContent".into(), s(""));
    o.properties.insert("parentNode".into(), Value::Null);
    o.properties.insert("firstChild".into(), Value::Null);
    o.properties.insert("lastChild".into(), Value::Null);
    o.properties.insert("nextSibling".into(), Value::Null);
    o.properties.insert("previousSibling".into(), Value::Null);
    o.properties.insert("firstElementChild".into(), Value::Null);
    o.properties.insert("lastElementChild".into(), Value::Null);
    o.properties.insert(
        "ownerDocument".into(),
        owner
            .map(|d| Value::Object(d.clone()))
            .unwrap_or(Value::Null),
    );
    Value::Object(vybe_runtime::heap::alloc(o))
}

fn find_by_local_name(node: &Arc<Mutex<Object>>, local: &str, out: &mut Vec<Value>) {
    let n = node.lock().unwrap();
    if matches!(n.properties.get("nodeType"), Some(Value::I32(t)) if *t == ELEMENT_NODE) {
        if let Some(Value::String(name)) = n.properties.get("localName") {
            if local == "*" || name.as_ref() == local {
                out.push(Value::Object(node.clone()));
            }
        }
    }
    let children = n.properties.get("childNodes").cloned();
    drop(n);
    if let Some(Value::Object(arr)) = children {
        let items: Vec<Value> = if let ObjectKind::Array(ref a) = arr.lock().unwrap().kind {
            a.clone()
        } else {
            return;
        };
        for child in items {
            if let Value::Object(c) = child {
                find_by_local_name(&c, local, out);
            }
        }
    }
}

#[cfg(test)]
mod html_grammar_tests {
    use super::*;
    use vybe_runtime::value::ObjectKind;

    /// A node's property, as a string.
    fn prop(node: &Value, name: &str) -> String {
        let Value::Object(o) = node else {
            return String::new();
        };
        match o.lock().unwrap().properties.get(name) {
            Some(Value::String(s)) => s.to_string(),
            Some(other) => format!("{other}"),
            None => String::new(),
        }
    }

    /// A node's array-valued property.
    fn list(node: &Value, name: &str) -> Vec<Value> {
        let Value::Object(o) = node else {
            return vec![];
        };
        let guard = o.lock().unwrap();
        match guard.properties.get(name) {
            Some(Value::Object(arr)) => match &arr.lock().unwrap().kind {
                ObjectKind::Array(elements) => elements.clone(),
                _ => vec![],
            },
            _ => vec![],
        }
    }

    fn children(node: &Value) -> Vec<Value> {
        list(node, "children")
    }

    /// Parse as HTML and hand back the root element.
    fn html(src: &str) -> Value {
        let doc = parse_markup(src, Grammar::Html).expect("HTML never fails to parse");
        let root = prop(&doc, "documentElement");
        assert!(!root.is_empty(), "a parsed document has a root element");
        let Value::Object(o) = &doc else {
            unreachable!()
        };
        let root = o
            .lock()
            .unwrap()
            .properties
            .get("documentElement")
            .cloned()
            .expect("documentElement");
        root
    }

    fn recoveries(src: &str) -> Vec<String> {
        let doc = parse_markup(src, Grammar::Html).expect("HTML never fails to parse");
        list(&doc, "__parseRecoveries")
            .iter()
            .map(|v| match v {
                Value::String(s) => s.to_string(),
                other => format!("{other}"),
            })
            .collect()
    }

    #[test]
    fn a_void_element_does_not_swallow_its_siblings() {
        // The rule that decides whether real HTML parses at all. Read as XML
        // this is a fatal error; read as HTML it is three children of <p>.
        let root = html("<p>one<br>two<img src=x>three</p>");
        assert_eq!(prop(&root, "tagName"), "p");
        let kids = children(&root);
        assert_eq!(
            kids.iter().map(|k| prop(k, "tagName")).collect::<Vec<_>>(),
            vec!["br", "img"],
            "the <br> and <img> are SIBLINGS, not a nest"
        );
        assert_eq!(
            prop(&root, "textContent"),
            "onetwothree",
            "and the text around them all belongs to the paragraph"
        );
    }

    #[test]
    fn the_same_markup_is_a_parse_error_as_xml() {
        // The other half of the claim: this is not a stricter XML reading,
        // it is a different grammar. Without a type argument nothing changes.
        assert!(
            parse_markup("<p>one<br>two</p>", Grammar::Xml).is_err(),
            "a bare <br> is a fatal XML error, which is why the type \
             argument had to start meaning something"
        );
    }

    #[test]
    fn an_omitted_end_tag_makes_siblings_not_a_nest() {
        let root = html("<ul><li>one<li>two<li>three</ul>");
        let items = children(&root);
        assert_eq!(items.len(), 3, "three list items, all siblings");
        for item in &items {
            assert_eq!(prop(item, "tagName"), "li");
            assert!(
                children(item).is_empty(),
                "an <li> must not contain the ones after it"
            );
        }
    }

    #[test]
    fn a_paragraph_is_closed_by_the_next_block() {
        let root = html("<div><p>one<p>two<ul><li>x</ul></div>");
        let kids = children(&root);
        assert_eq!(
            kids.iter().map(|k| prop(k, "tagName")).collect::<Vec<_>>(),
            vec!["p", "p", "ul"],
            "both paragraphs and the list are siblings inside the div"
        );
    }

    #[test]
    fn tag_and_attribute_names_fold_to_lowercase() {
        // Everything downstream compares against lowercase literals —
        // VOID_ELEMENTS here, `control_kind` and `ua::declarations_for` in
        // the toolkit. Folding at the parse is what makes those hit.
        let root = html("<DIV CLASS='x'><BR></DIV>");
        assert_eq!(prop(&root, "tagName"), "div");
        assert_eq!(prop(&root, "className"), "x");
        assert_eq!(
            children(&root)
                .iter()
                .map(|k| prop(k, "tagName"))
                .collect::<Vec<_>>(),
            vec!["br"],
            "an uppercase <BR> is still void"
        );
    }

    #[test]
    fn common_named_entities_decode_and_an_unknown_one_survives() {
        assert_eq!(
            prop(&html("<p>a&nbsp;b&mdash;c</p>"), "textContent"),
            "a\u{a0}b—c"
        );
        // An unlisted name is left as written rather than deleting the run
        // it sits in — wrong, but visible.
        assert!(
            prop(&html("<p>x&fnof;y</p>"), "textContent").contains('x'),
            "an unknown entity must not eat its text run"
        );
    }

    #[test]
    fn the_five_xml_predefined_entities_still_decode_in_html() {
        // `unescape_with` routes EVERY named reference to the resolver,
        // including the five XML predefines — so leaving them out of the
        // table made `&amp;` read back as the literal `&amp;` in HTML while
        // still decoding in XML. Found by probing, not by reasoning.
        assert_eq!(prop(&html("<p>a &amp; b</p>"), "textContent"), "a & b");
        assert_eq!(prop(&html("<p>&lt;tag&gt;</p>"), "textContent"), "<tag>");
    }

    #[test]
    fn raw_text_is_not_read_as_markup() {
        // Measured before the fix: this truncated to `"if (a "` — no error,
        // no recovery, the rest of the script swallowed by an element named
        // `b)`. The silent-wrong-answer case the recovery list exists for,
        // which is exactly why it could not be left to the recovery list.
        let root = html("<script>if (a < b) { x() }</script>");
        assert_eq!(prop(&root, "tagName"), "script");
        assert_eq!(
            prop(&root, "textContent"),
            "if (a < b) { x() }",
            "a script's text survives its own comparisons"
        );

        let style = html("<style>p { color: red }</style>");
        assert_eq!(prop(&style, "textContent"), "p { color: red }");
    }

    #[test]
    fn recovery_is_reported_because_a_tolerant_parser_stops_erroring() {
        // The signal that replaces `<parsererror>`. Once malformed input
        // parses, "it worked" and "I repaired it" look identical from the
        // outside unless the parser says so.
        assert!(
            recoveries("<p>clean</p>").is_empty(),
            "well-formed HTML reports no repairs"
        );
        let repaired = recoveries("<ul><li>one<li>two</ul>");
        assert!(
            repaired.iter().any(|r| r.contains("implied </li>")),
            "the implied end tag is reported, got {repaired:?}"
        );
        let stray = recoveries("<p>x</p></div>");
        assert!(
            stray.iter().any(|r| r.contains("stray </div>")),
            "an unmatched end tag is reported, got {stray:?}"
        );
    }

    /// The XML path's text extraction, asserted directly.
    ///
    /// The reader loop was lifted out into `drive` and a sink; every text run
    /// now arrives as `TreeSink::text` instead of being pushed inline. This is
    /// the one behaviour that refactor could have changed silently — a tree
    /// with the right shape and no content reads as a downstream bug — so it
    /// is checked at the parser rather than inferred from a suite.
    #[test]
    fn an_xml_text_run_becomes_a_text_node_with_its_data() {
        let doc = parse_markup("<root><title>Post</title></root>", Grammar::Xml)
            .expect("well-formed XML parses");
        let Value::Object(o) = &doc else {
            unreachable!("a parse answers a document object")
        };
        let root = o
            .lock()
            .unwrap()
            .properties
            .get("documentElement")
            .cloned()
            .expect("documentElement");
        assert_eq!(prop(&root, "tagName"), "root");
        assert_eq!(prop(&root, "textContent"), "Post");
        let title = children(&root).remove(0);
        assert_eq!(prop(&title, "tagName"), "title");
        let text = list(&title, "childNodes");
        assert_eq!(text.len(), 1, "one text node");
        assert_eq!(prop(&text[0], "nodeType"), "3");
        assert_eq!(prop(&text[0], "nodeValue"), "Post");
        assert_eq!(prop(&text[0], "nodeName"), "#text");
    }

    #[test]
    fn an_end_tag_closes_the_nearest_matching_ancestor() {
        // `<b><i>x</b>` — the bold ends and takes the italic with it. A
        // browser reopens the italic afterwards (the adoption agency
        // algorithm), which is out of scope; this asserts what we DO, and
        // that we said so.
        let root = html("<b><i>x</b>y");
        assert_eq!(prop(&root, "tagName"), "b");
        let repaired = recoveries("<b><i>x</b>y");
        assert!(
            repaired.iter().any(|r| r.contains("still-open i")),
            "the divergence is recorded, got {repaired:?}"
        );
    }
}

/// `parseFromString(…, "text/html")` against the REAL document.
///
/// These assert through the toolkit rather than through the return value,
/// because the return value is now a handle and the tree is the answer. What
/// each one is really checking is that a parsed page is not a second kind of
/// document: it cascades, it lays out, it serialises, and every `web:dom`
/// operation is about it.
#[cfg(all(test, feature = "gui"))]
mod html_document_tests {
    use super::*;
    use vybe_widgets::dom;

    /// Parse, and hand back the document the handle names.
    fn parse(source: &str) -> dom::DocumentId {
        crate::engine_widgets::install();
        let handle = parse_html_document(source);
        let Value::Object(o) = &handle else {
            panic!("parseFromString answers a handle");
        };
        let id = o
            .lock()
            .unwrap()
            .properties
            .get("__document")
            .map(|v| v.as_f64() as dom::DocumentId)
            .expect("the handle names its document");
        assert_ne!(id, 0, "a parsed document is a real, addressable document");
        id
    }

    fn html_of(document: dom::DocumentId) -> String {
        dom::with_document(document, |doc| doc.to_html()).unwrap_or_default()
    }

    #[test]
    fn a_parsed_page_is_a_document_the_engine_can_answer_about() {
        // The whole point in one assertion: the tree came from markup and
        // `getElementById` — an ordinary `web:dom` call — finds it. Before
        // this, a parsed tree and a live document were different objects and
        // no document operation reached the parsed one.
        let document = parse("<div id='wrap'><p>hello</p></div>");
        let found = dom::with_document(document, |doc| doc.get_element_by_id("wrap")).flatten();
        assert!(found.is_some(), "the parsed element is in the document");
        let html = html_of(document);
        assert!(html.contains("<div id=\"wrap\">"), "got {html}");
        assert!(html.contains("<p>"), "got {html}");
    }

    #[test]
    fn a_style_element_becomes_the_author_stylesheet() {
        // A `<style>` is not decoration on the way past: its text IS the
        // cascade's author origin. A parsed page whose rules did not apply
        // would be a document in shape only.
        let document = parse("<html><head><style>p { color: #ff0000 }</style></head><body><p>x</p></body></html>");
        let colour = dom::with_document(document, |doc| {
            let p = doc.query_selector("p").expect("the paragraph is in the tree");
            doc.get_computed_style(p).color
        })
        .expect("the document is open");
        assert_eq!(colour, Some(0xffff0000), "the parsed rule reached the cascade");
    }

    #[test]
    fn an_inline_style_attribute_is_a_declaration_block() {
        // `style=""` is the last origin in the cascade, and it arrived as an
        // inert attribute for as long as the parser has existed.
        let document = parse("<p style='color: #00ff00'>x</p>");
        let colour = dom::with_document(document, |doc| {
            let p = doc.query_selector("p").expect("the paragraph is in the tree");
            doc.get_computed_style(p).color
        })
        .expect("the document is open");
        assert_eq!(colour, Some(0xff00ff00));
    }

    #[test]
    fn an_inline_style_beats_a_rule_that_selects_the_same_element() {
        // Two origins, one element — which is the only way to show the
        // stylesheet and the attribute are genuinely the same cascade rather
        // than two writes racing.
        let document = parse(
            "<style>p { color: #ff0000 }</style><p style='color: #0000ff'>x</p>",
        );
        let colour = dom::with_document(document, |doc| {
            let p = doc.query_selector("p").expect("the paragraph is in the tree");
            doc.get_computed_style(p).color
        })
        .expect("the document is open");
        assert_eq!(colour, Some(0xff0000ff), "inline wins");
    }

    #[test]
    fn a_comment_survives_the_parse_and_serialises_back() {
        // The alternative to a comment node is dropping one, which is a
        // silent difference between the markup handed in and the tree handed
        // back — and it would take `<![CDATA[…]]>` and `<?…?>` with it.
        let document = parse("<div><!-- note --><p>x</p></div>");
        let html = html_of(document);
        assert!(html.contains("<!-- note -->"), "got {html}");
    }

    #[test]
    fn a_comment_is_not_part_of_its_parents_text() {
        let document = parse("<div><!-- hidden --><p>shown</p></div>");
        let text = dom::with_document(document, |doc| {
            let div = doc.query_selector("div").expect("the div is in the tree");
            doc.text_content(div)
        })
        .expect("the document is open");
        assert!(!text.contains("hidden"), "got {text:?}");
        assert!(text.contains("shown"), "got {text:?}");
    }

    #[test]
    fn a_title_names_the_document() {
        let document = parse("<html><head><title>Report</title></head><body></body></html>");
        let title = dom::with_document(document, |doc| doc.title()).unwrap_or_default();
        assert_eq!(title, "Report");
    }

    #[test]
    fn text_inside_a_leaf_element_is_kept_as_its_text() {
        // `<span>`, `<option>`, `<td>` and `<label>` are leaves here: they
        // refuse a text NODE, and for a leaf its text and its `textContent`
        // are the same fact. Falling through without this dropped the content
        // of every one of them.
        let document = parse("<p>before <span>inside</span> after</p>");
        let text = dom::with_document(document, |doc| {
            let span = doc.query_selector("span").expect("the span is in the tree");
            doc.text_content(span)
        })
        .expect("the document is open");
        assert_eq!(text, "inside");
    }

    #[test]
    fn the_page_structure_elements_hold_their_children() {
        // `<html>`, `<head>` and `<body>` are the three elements nothing ever
        // created until a parser did. As leaves they refused every child, so
        // a whole page arrived empty.
        let document = parse("<html><head><meta charset='utf-8'></head><body><p>x</p></body></html>");
        let (has_meta, has_p) = dom::with_document(document, |doc| {
            (
                doc.query_selector("head meta").is_some(),
                doc.query_selector("body p").is_some(),
            )
        })
        .expect("the document is open");
        assert!(has_meta, "the metadata is inside the head");
        assert!(has_p, "the paragraph is inside the body");
    }

    #[test]
    fn a_parsed_document_is_not_the_page() {
        // `parseFromString` must not touch the browsing context's own
        // document — that is what makes it usable for reading a fragment.
        crate::engine_widgets::install();
        let page = crate::html::active_document();
        let parsed = parse("<p id='only-here'>x</p>");
        assert_ne!(parsed, page);
        let leaked = dom::with_document(page, |doc| doc.get_element_by_id("only-here")).flatten();
        assert!(leaked.is_none(), "the parse stayed in its own document");
    }
}
