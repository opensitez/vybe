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
use vybe_bytecode::value::{Object, ObjectKind};
use vybe_bytecode::{Value, VM};

use quick_xml::events::Event;
use quick_xml::Reader;

// ── nodeType constants (WHATWG DOM Living Standard §4.4) ──────────
const ELEMENT_NODE: i32 = 1;
const TEXT_NODE: i32 = 3;
const CDATA_SECTION_NODE: i32 = 4;
const PROCESSING_INSTRUCTION_NODE: i32 = 7;
const COMMENT_NODE: i32 = 8;
const DOCUMENT_NODE: i32 = 9;

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
    vm.register_host_fn("web:dom-parser", "parserNew", Box::new(|_ctx, _args| {
        let mut obj = Object::new();
        obj.properties.insert("__type".into(), s("DOMParser"));
        Value::Object(Arc::new(Mutex::new(obj)))
    }));

    vm.register_host_fn("web:dom-parser", "parseFromString", Box::new(|_ctx, args| {
        // Spec: `DOMParser.parseFromString(string, type)`. We accept
        // both the instance-call shape (args[0] = DOMParser, args[1]
        // = string, args[2] = type) and the flat shorthand
        // (args[0] = string, args[1] = type) so callers that don't
        // construct a parser still work.
        let (xml_arg, _type_arg) = match args.first() {
            Some(Value::Object(o)) if o.lock().unwrap().properties.get("__type")
                .and_then(|v| match v { Value::String(s) => Some(s.as_ref().to_string()), _ => None })
                .as_deref() == Some("DOMParser") =>
            {
                (args.get(1).cloned().unwrap_or(Value::Null),
                 args.get(2).cloned().unwrap_or(Value::Null))
            }
            _ => (args.first().cloned().unwrap_or(Value::Null),
                  args.get(1).cloned().unwrap_or(Value::Null)),
        };
        let xml = match xml_arg {
            Value::String(s) => s.to_string(),
            other => format!("{}", other),
        };
        match parse_xml(&xml) {
            Ok(doc) => doc,
            Err(_) => parse_error_document(&xml),
        }
    }));

    // Convenience flat fn: `parse(s)` shorthand callers use without an
    // explicit DOMParser instance. Same return as `parseFromString`.
    vm.register_host_fn("web:dom-parser", "parse", Box::new(|_ctx, args| {
        let xml = match args.first() {
            Some(Value::String(s)) => s.to_string(),
            Some(other) => format!("{}", other),
            None => String::new(),
        };
        match parse_xml(&xml) {
            Ok(doc) => doc,
            Err(_) => parse_error_document(&xml),
        }
    }));

    // ── XMLSerializer ────────────────────────────────────────────
    vm.register_host_fn("web:dom-parser", "serializerNew", Box::new(|_ctx, _args| {
        let mut obj = Object::new();
        obj.properties.insert("__type".into(), s("XMLSerializer"));
        Value::Object(Arc::new(Mutex::new(obj)))
    }));

    vm.register_host_fn("web:dom-parser", "serializeToString", Box::new(|_ctx, args| {
        // Accepts both `(serializer, node)` and `(node)` shapes.
        let node = match args.first() {
            Some(Value::Object(o)) if o.lock().unwrap().properties.get("__type")
                .and_then(|v| match v { Value::String(s) => Some(s.as_ref().to_string()), _ => None })
                .as_deref() == Some("XMLSerializer") =>
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
    }));

    // Convenience: `toString(node)` matches the legacy shorthand.
    vm.register_host_fn("web:dom-parser", "toString", Box::new(|_ctx, args| {
        let mut out = String::new();
        if let Some(Value::Object(o)) = args.first() {
            serialize_node(o, &mut out);
        }
        s(&out)
    }));

    // Convenience helper retained from the pre-spec API: `load(path)`
    // reads from disk and parses. Real WHATWG flow is `fetch(url)` →
    // `.text()` → `parseFromString` — provided here for parity with
    // VB `XDocument.Load(path)` test patterns.
    vm.register_host_fn("web:dom-parser", "load", Box::new(|_ctx, args| {
        let path = match args.first() {
            Some(Value::String(s)) => s.to_string(),
            Some(other) => format!("{}", other),
            None => return Value::Null,
        };
        match std::fs::read_to_string(&path) {
            Ok(xml) => parse_xml(&xml).unwrap_or_else(|_| parse_error_document(&xml)),
            Err(_) => Value::Null,
        }
    }));

    // ── Node / Element / Document method helpers ─────────────────
    // Properties (`tagName`, `nodeType`, `childNodes`, etc.) are set
    // directly on the object during parse, so user code reads them
    // via plain `Op::STRUCT_GET` — no host fn round-trip. The host
    // fns below cover the *computed* methods that walk the tree.

    vm.register_host_fn("web:dom-parser", "getElementById", Box::new(|_ctx, args| {
        let Some(Value::Object(root)) = args.first() else { return Value::Null; };
        let target = match args.get(1) {
            Some(Value::String(s)) => s.to_string(),
            Some(other) => format!("{}", other),
            None => return Value::Null,
        };
        find_by_id(root, &target).unwrap_or(Value::Null)
    }));

    vm.register_host_fn("web:dom-parser", "getElementsByTagName", Box::new(|_ctx, args| {
        let Some(Value::Object(root)) = args.first() else { return empty_array(); };
        let tag = match args.get(1) {
            Some(Value::String(s)) => s.to_string(),
            Some(other) => format!("{}", other),
            None => return empty_array(),
        };
        let mut out: Vec<Value> = Vec::new();
        find_by_tag_name(root, &tag, &mut out);
        make_array(out)
    }));

    vm.register_host_fn("web:dom-parser", "getElementsByClassName", Box::new(|_ctx, args| {
        let Some(Value::Object(root)) = args.first() else { return empty_array(); };
        let cls = match args.get(1) {
            Some(Value::String(s)) => s.to_string(),
            Some(other) => format!("{}", other),
            None => return empty_array(),
        };
        let mut out: Vec<Value> = Vec::new();
        find_by_class_name(root, &cls, &mut out);
        make_array(out)
    }));

    vm.register_host_fn("web:dom-parser", "getAttribute", Box::new(|_ctx, args| {
        let Some(Value::Object(elem)) = args.first() else { return Value::Null; };
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
    }));

    vm.register_host_fn("web:dom-parser", "hasAttribute", Box::new(|_ctx, args| {
        let Some(Value::Object(elem)) = args.first() else { return Value::Bool(false); };
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
    }));

    // ── Selectors API Level 1 (W3C) ──────────────────────────────
    // CSS selector subset: tag, `*`, `#id`, `.class`, `[attr]` /
    // `[attr="val"]` / `[attr*="val"]` / `[attr^="val"]` /
    // `[attr$="val"]`, descendant (` `), child (`>`), adjacent sibling
    // (`+`), general sibling (`~`), compound (`tag.class#id[attr]`),
    // and selector lists (comma-separated). Pseudo-classes
    // (`:first-child`, `:nth-child`) are Phase 2.5 follow-up.

    vm.register_host_fn("web:dom-parser", "querySelector", Box::new(|_ctx, args| {
        let Some(Value::Object(root)) = args.first() else { return Value::Null; };
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
    }));

    vm.register_host_fn("web:dom-parser", "querySelectorAll", Box::new(|_ctx, args| {
        let Some(Value::Object(root)) = args.first() else { return empty_array(); };
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
    }));

    vm.register_host_fn("web:dom-parser", "matches", Box::new(|_ctx, args| {
        let Some(Value::Object(elem)) = args.first() else { return Value::Bool(false); };
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
    }));

    // ── DOM Mutation (Phase 3) ───────────────────────────────────
    // WHATWG DOM Living Standard §4.5: Document factories +
    // Node mutation methods. Newly-created nodes start orphaned
    // (parentNode=null, no siblings, ownerDocument set when known).

    vm.register_host_fn("web:dom-parser", "createElement", Box::new(|_ctx, args| {
        let tag = match args.get(1).or(args.first()) {
            Some(Value::String(s)) => s.to_string(),
            Some(other) => format!("{}", other),
            None => return Value::Null,
        };
        // Detect (doc, tag) shape vs (tag) shape: if first arg is a
        // Document, owner is set; otherwise orphan.
        let owner: Option<Arc<Mutex<Object>>> = match args.first() {
            Some(Value::Object(o)) => {
                let lock = o.lock().unwrap();
                let is_doc = matches!(lock.properties.get("nodeType"), Some(Value::I32(t)) if *t == DOCUMENT_NODE);
                if is_doc { Some(o.clone()) } else { None }
            }
            _ => None,
        };
        new_element_node(&tag, owner.as_ref())
    }));

    vm.register_host_fn("web:dom-parser", "createTextNode", Box::new(|_ctx, args| {
        let text = match args.get(1).or(args.first()) {
            Some(Value::String(s)) => s.to_string(),
            Some(other) => format!("{}", other),
            None => String::new(),
        };
        let owner: Option<Arc<Mutex<Object>>> = match args.first() {
            Some(Value::Object(o)) => {
                let lock = o.lock().unwrap();
                let is_doc = matches!(lock.properties.get("nodeType"), Some(Value::I32(t)) if *t == DOCUMENT_NODE);
                if is_doc { Some(o.clone()) } else { None }
            }
            _ => None,
        };
        let node = make_node(TEXT_NODE, "#text", Some(&text));
        if let (Value::Object(n), Some(d)) = (&node, &owner) {
            n.lock().unwrap().properties.insert("ownerDocument".into(), Value::Object(d.clone()));
        }
        if let Value::Object(n) = &node {
            n.lock().unwrap().properties.insert("textContent".into(), s(&text));
        }
        node
    }));

    vm.register_host_fn("web:dom-parser", "createComment", Box::new(|_ctx, args| {
        let text = match args.get(1).or(args.first()) {
            Some(Value::String(s)) => s.to_string(),
            Some(other) => format!("{}", other),
            None => String::new(),
        };
        let node = make_node(COMMENT_NODE, "#comment", Some(&text));
        if let Value::Object(n) = &node {
            n.lock().unwrap().properties.insert("textContent".into(), s(&text));
        }
        node
    }));

    vm.register_host_fn("web:dom-parser", "setAttribute", Box::new(|_ctx, args| {
        let Some(Value::Object(elem)) = args.first() else { return Value::Null; };
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
    }));

    vm.register_host_fn("web:dom-parser", "removeAttribute", Box::new(|_ctx, args| {
        let Some(Value::Object(elem)) = args.first() else { return Value::Null; };
        let name = match args.get(1) {
            Some(Value::String(s)) => s.to_string(),
            Some(other) => format!("{}", other),
            None => return Value::Null,
        };
        let elem_lock = elem.lock().unwrap();
        let attrs = elem_lock.properties.get("attributes").cloned();
        drop(elem_lock);
        if let Some(Value::Object(a)) = attrs {
            a.lock().unwrap().properties.remove(&name);
        }
        let mut elem_w = elem.lock().unwrap();
        if name == "id" {
            elem_w.properties.insert("id".into(), s(""));
        } else if name == "class" {
            elem_w.properties.insert("className".into(), s(""));
        }
        Value::Null
    }));

    vm.register_host_fn("web:dom-parser", "appendChild", Box::new(|_ctx, args| {
        let Some(Value::Object(parent)) = args.first() else { return Value::Null; };
        let Some(Value::Object(child)) = args.get(1) else { return Value::Null; };
        // Detach child from any current parent first (DOM semantics:
        // appendChild is a move, not a copy — `pre-insert` step in spec).
        detach_from_parent(child);
        // Append + rewire siblings.
        append_child_inner(parent, child);
        Value::Object(child.clone())
    }));

    vm.register_host_fn("web:dom-parser", "removeChild", Box::new(|_ctx, args| {
        let Some(Value::Object(parent)) = args.first() else { return Value::Null; };
        let Some(Value::Object(child)) = args.get(1) else { return Value::Null; };
        if remove_child_inner(parent, child) {
            Value::Object(child.clone())
        } else {
            Value::Null
        }
    }));

    vm.register_host_fn("web:dom-parser", "insertBefore", Box::new(|_ctx, args| {
        let Some(Value::Object(parent)) = args.first() else { return Value::Null; };
        let Some(Value::Object(new_child)) = args.get(1) else { return Value::Null; };
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
    }));

    vm.register_host_fn("web:dom-parser", "replaceChild", Box::new(|_ctx, args| {
        let Some(Value::Object(parent)) = args.first() else { return Value::Null; };
        let Some(Value::Object(new_child)) = args.get(1) else { return Value::Null; };
        let Some(Value::Object(old_child)) = args.get(2) else { return Value::Null; };
        detach_from_parent(new_child);
        if !insert_before_inner(parent, new_child, old_child) {
            return Value::Null;
        }
        let _ = remove_child_inner(parent, old_child);
        Value::Object(old_child.clone())
    }));

    vm.register_host_fn("web:dom-parser", "cloneNode", Box::new(|_ctx, args| {
        let Some(Value::Object(node)) = args.first() else { return Value::Null; };
        let deep = matches!(args.get(1), Some(Value::Bool(true)));
        clone_node(node, deep)
    }));

    // ── DocumentFragment + Namespace-aware variants (Phase 4) ────
    vm.register_host_fn("web:dom-parser", "createDocumentFragment", Box::new(|_ctx, args| {
        let owner: Option<Arc<Mutex<Object>>> = match args.first() {
            Some(Value::Object(o)) => {
                let lock = o.lock().unwrap();
                let is_doc = matches!(lock.properties.get("nodeType"), Some(Value::I32(t)) if *t == DOCUMENT_NODE);
                if is_doc { Some(o.clone()) } else { None }
            }
            _ => None,
        };
        new_document_fragment(owner.as_ref())
    }));

    vm.register_host_fn("web:dom-parser", "createElementNS", Box::new(|_ctx, args| {
        let owner: Option<Arc<Mutex<Object>>> = match args.first() {
            Some(Value::Object(o)) => {
                let lock = o.lock().unwrap();
                let is_doc = matches!(lock.properties.get("nodeType"), Some(Value::I32(t)) if *t == DOCUMENT_NODE);
                if is_doc { Some(o.clone()) } else { None }
            }
            _ => None,
        };
        let (ns_arg, qname_arg) = if owner.is_some() {
            (args.get(1), args.get(2))
        } else {
            (args.first(), args.get(1))
        };
        let ns = match ns_arg {
            Some(Value::String(s)) => Some(s.to_string()),
            Some(Value::Null) | None => None,
            Some(other) => Some(format!("{}", other)),
        };
        let qname = match qname_arg {
            Some(Value::String(s)) => s.to_string(),
            Some(other) => format!("{}", other),
            None => return Value::Null,
        };
        let elem = new_element_node(&qname, owner.as_ref());
        if let Value::Object(e) = &elem {
            e.lock().unwrap().properties.insert("namespaceURI".into(),
                ns.map(|n| s(&n)).unwrap_or(Value::Null));
        }
        elem
    }));

    vm.register_host_fn("web:dom-parser", "setAttributeNS", Box::new(|_ctx, args| {
        let Some(Value::Object(elem)) = args.first() else { return Value::Null; };
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
            elem.lock().unwrap().properties.insert("className".into(), s(&val));
        }
        Value::Null
    }));

    vm.register_host_fn("web:dom-parser", "getAttributeNS", Box::new(|_ctx, args| {
        let Some(Value::Object(elem)) = args.first() else { return Value::Null; };
        let name = match args.get(2).or(args.get(1)) {
            Some(Value::String(s)) => s.to_string(),
            Some(other) => format!("{}", other),
            None => return Value::Null,
        };
        let attrs = elem.lock().unwrap().properties.get("attributes").cloned();
        if let Some(Value::Object(a)) = attrs {
            return a.lock().unwrap().properties.get(&name).cloned().unwrap_or(Value::Null);
        }
        Value::Null
    }));

    vm.register_host_fn("web:dom-parser", "hasAttributeNS", Box::new(|_ctx, args| {
        let Some(Value::Object(elem)) = args.first() else { return Value::Bool(false); };
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
    }));

    vm.register_host_fn("web:dom-parser", "removeAttributeNS", Box::new(|_ctx, args| {
        let Some(Value::Object(elem)) = args.first() else { return Value::Null; };
        let name = match args.get(2).or(args.get(1)) {
            Some(Value::String(s)) => s.to_string(),
            Some(other) => format!("{}", other),
            None => return Value::Null,
        };
        let attrs = elem.lock().unwrap().properties.get("attributes").cloned();
        if let Some(Value::Object(a)) = attrs {
            a.lock().unwrap().properties.remove(&name);
        }
        Value::Null
    }));

    vm.register_host_fn("web:dom-parser", "getElementsByTagNameNS", Box::new(|_ctx, args| {
        let Some(Value::Object(root)) = args.first() else { return empty_array(); };
        let local = match args.get(2).or(args.get(1)) {
            Some(Value::String(s)) => s.to_string(),
            Some(other) => format!("{}", other),
            None => return empty_array(),
        };
        let mut out: Vec<Value> = Vec::new();
        find_by_local_name(root, &local, &mut out);
        make_array(out)
    }));

    vm.register_host_fn("web:dom-parser", "closest", Box::new(|_ctx, args| {
        let Some(Value::Object(elem)) = args.first() else { return Value::Null; };
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
    }));
}

// ── Parser entry point ────────────────────────────────────────────

fn parse_xml(xml: &str) -> Result<Value, String> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(false);
    reader.config_mut().expand_empty_elements = false;

    // Build the Document node first; its childNodes will be populated
    // as we read events.
    let document_obj = make_node(DOCUMENT_NODE, "#document", None);

    // Use a stack of parent contexts to handle nesting. Each entry is
    // an Arc<Mutex<Object>> (the parent node) and the in-progress
    // childNodes Vec (Element-style siblings).
    let mut node_stack: Vec<Arc<Mutex<Object>>> = vec![match &document_obj {
        Value::Object(o) => o.clone(),
        _ => unreachable!(),
    }];

    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let elem = make_element_from_start(&e);
                push_child(&node_stack, elem.clone());
                if let Value::Object(o) = elem {
                    node_stack.push(o);
                }
            }
            Ok(Event::Empty(e)) => {
                let elem = make_element_from_start(&e);
                push_child(&node_stack, elem);
                // Empty/self-closing — no push to stack; siblings
                // continue at the current parent.
            }
            Ok(Event::End(_)) => {
                if node_stack.len() > 1 {
                    node_stack.pop();
                }
            }
            Ok(Event::Text(t)) => {
                let text = t.unescape().map(|c| c.into_owned()).unwrap_or_default();
                if !text.is_empty() {
                    let node = make_node(TEXT_NODE, "#text", Some(&text));
                    push_child(&node_stack, node);
                }
            }
            Ok(Event::CData(c)) => {
                let bytes = c.into_inner().into_owned();
                let text = String::from_utf8_lossy(&bytes).into_owned();
                let node = make_node(CDATA_SECTION_NODE, "#cdata-section", Some(&text));
                push_child(&node_stack, node);
            }
            Ok(Event::Comment(c)) => {
                let text = c.unescape().map(|c| c.into_owned()).unwrap_or_default();
                let node = make_node(COMMENT_NODE, "#comment", Some(&text));
                push_child(&node_stack, node);
            }
            Ok(Event::PI(pi)) => {
                let raw = String::from_utf8_lossy(pi.as_ref()).into_owned();
                // Spec: target = first whitespace-delimited token; data
                // = the rest. `<?xml-stylesheet href="x"?>` → target
                // "xml-stylesheet", data `href="x"`.
                let mut parts = raw.splitn(2, char::is_whitespace);
                let target = parts.next().unwrap_or("").to_string();
                let data = parts.next().unwrap_or("").trim().to_string();
                let node = make_pi_node(&target, &data);
                push_child(&node_stack, node);
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

    Ok(document_obj)
}

fn parse_error_document(xml: &str) -> Value {
    // WHATWG: parseerror documents are still Documents whose root
    // element is `<parsererror>`. We mirror that minimally.
    let doc = make_node(DOCUMENT_NODE, "#document", None);
    let err_elem = {
        let mut o = Object::new();
        o.properties.insert("__type".into(), s("Element"));
        o.properties.insert("nodeType".into(), Value::I32(ELEMENT_NODE));
        o.properties.insert("nodeName".into(), s("parsererror"));
        o.properties.insert("tagName".into(), s("parsererror"));
        o.properties.insert("localName".into(), s("parsererror"));
        o.properties.insert("prefix".into(), Value::Null);
        o.properties.insert("namespaceURI".into(), Value::Null);
        o.properties.insert("nodeValue".into(), Value::Null);
        o.properties.insert("attributes".into(), make_empty_object());
        o.properties.insert("childNodes".into(), make_array(vec![]));
        o.properties.insert("children".into(), make_array(vec![]));
        o.properties.insert("textContent".into(), s(xml));
        o.properties.insert("id".into(), s(""));
        o.properties.insert("className".into(), s(""));
        Value::Object(Arc::new(Mutex::new(o)))
    };
    if let Value::Object(d) = &doc {
        if let Value::Object(arr) = d.lock().unwrap().properties.get("childNodes").cloned().unwrap_or(Value::Null) {
            if let ObjectKind::Array(ref mut items) = arr.lock().unwrap().kind {
                items.push(err_elem.clone());
            }
        }
        if let Value::Object(d_obj) = &doc {
            d_obj.lock().unwrap().properties.insert("documentElement".into(), err_elem);
        }
    }
    if let Value::Object(d) = &doc {
        finalize_node_tree(d, None, Some(d));
    }
    doc
}

// ── Element construction ──────────────────────────────────────────

fn make_element_from_start(e: &quick_xml::events::BytesStart) -> Value {
    let raw_name = String::from_utf8_lossy(e.name().as_ref()).into_owned();
    // Split optional prefix per Namespaces in XML 1.0 §3.
    let (prefix, local_name) = match raw_name.split_once(':') {
        Some((p, ln)) => (Some(p.to_string()), ln.to_string()),
        None => (None, raw_name.clone()),
    };
    let mut o = Object::new();
    o.type_id = type_id_for(ELEMENT_NODE);
    o.properties.insert("__type".into(), s("Element"));
    o.properties.insert("nodeType".into(), Value::I32(ELEMENT_NODE));
    o.properties.insert("nodeName".into(), s(&raw_name));
    o.properties.insert("tagName".into(), s(&raw_name));
    o.properties.insert("localName".into(), s(&local_name));
    o.properties.insert("prefix".into(), prefix.map(|p| s(&p)).unwrap_or(Value::Null));
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
        let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
        let val = attr.unescape_value().map(|c| c.into_owned()).unwrap_or_default();
        if key == "id" {
            id_val = val.clone();
        } else if key == "class" {
            class_val = val.clone();
        }
        attrs.properties.insert(key, s(&val));
    }
    o.properties.insert("attributes".into(), Value::Object(Arc::new(Mutex::new(attrs))));
    o.properties.insert("id".into(), s(&id_val));
    o.properties.insert("className".into(), s(&class_val));

    // childNodes / children populated later as parse progresses.
    o.properties.insert("childNodes".into(), make_array(vec![]));
    o.properties.insert("children".into(), make_array(vec![]));
    Value::Object(Arc::new(Mutex::new(o)))
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
    o.properties.insert("nodeType".into(), Value::I32(node_type));
    o.properties.insert("nodeName".into(), s(node_name));
    o.properties.insert("nodeValue".into(), match value {
        Some(v) => s(v),
        None => Value::Null,
    });
    if node_type == DOCUMENT_NODE {
        o.properties.insert("childNodes".into(), make_array(vec![]));
        o.properties.insert("documentElement".into(), Value::Null);
        o.properties.insert("doctype".into(), Value::Null);
    }
    Value::Object(Arc::new(Mutex::new(o)))
}

fn make_pi_node(target: &str, data: &str) -> Value {
    let mut o = Object::new();
    o.type_id = type_id_for(PROCESSING_INSTRUCTION_NODE);
    o.properties.insert("__type".into(), s("ProcessingInstruction"));
    o.properties.insert("nodeType".into(), Value::I32(PROCESSING_INSTRUCTION_NODE));
    o.properties.insert("nodeName".into(), s(target));
    o.properties.insert("target".into(), s(target));
    o.properties.insert("data".into(), s(data));
    o.properties.insert("nodeValue".into(), s(data));
    Value::Object(Arc::new(Mutex::new(o)))
}

fn push_child(stack: &[Arc<Mutex<Object>>], child: Value) {
    let parent = stack.last().expect("dom parser: empty stack");
    let parent_lock = parent.lock().unwrap();
    if let Some(Value::Object(arr)) = parent_lock.properties.get("childNodes") {
        let arr = arr.clone();
        drop(parent_lock);
        if let ObjectKind::Array(ref mut items) = arr.lock().unwrap().kind {
            items.push(child);
        }
    }
}

// ── Post-walk: parentNode / ownerDocument / textContent / siblings ──

fn finalize_node_tree(
    node: &Arc<Mutex<Object>>,
    parent: Option<&Arc<Mutex<Object>>>,
    document: Option<&Arc<Mutex<Object>>>,
) {
    {
        let mut n = node.lock().unwrap();
        n.properties.insert("parentNode".into(),
            parent.map(|p| Value::Object(p.clone())).unwrap_or(Value::Null));
        n.properties.insert("ownerDocument".into(),
            match document {
                Some(d) if !Arc::ptr_eq(d, node) => Value::Object(d.clone()),
                _ => Value::Null,
            });
    }
    let children = {
        let n = node.lock().unwrap();
        n.properties.get("childNodes").cloned()
    };
    let Some(Value::Object(arr)) = children else { return };
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
            let prev = if i > 0 { child_values.get(i - 1).cloned() } else { None };
            let next = child_values.get(i + 1).cloned();
            let mut c = child_obj.lock().unwrap();
            c.properties.insert("previousSibling".into(), prev.unwrap_or(Value::Null));
            c.properties.insert("nextSibling".into(), next.unwrap_or(Value::Null));
        }
    }

    // Element-only children list + firstElementChild / lastElementChild.
    let element_children: Vec<Value> = child_values.iter()
        .filter(|v| matches!(v, Value::Object(o)
            if matches!(o.lock().unwrap().properties.get("nodeType"), Some(Value::I32(1)))))
        .cloned()
        .collect();
    let first_elem = element_children.first().cloned().unwrap_or(Value::Null);
    let last_elem = element_children.last().cloned().unwrap_or(Value::Null);

    let mut n = node.lock().unwrap();
    n.properties.insert("firstChild".into(),
        child_values.first().cloned().unwrap_or(Value::Null));
    n.properties.insert("lastChild".into(),
        child_values.last().cloned().unwrap_or(Value::Null));

    let is_element_or_doc = matches!(n.properties.get("nodeType"), Some(Value::I32(node_type))
        if *node_type == ELEMENT_NODE || *node_type == DOCUMENT_NODE);
    let is_text_like = matches!(n.properties.get("nodeType"), Some(Value::I32(t))
        if *t == TEXT_NODE || *t == CDATA_SECTION_NODE);

    if is_element_or_doc {
        // Replace children Array contents in place.
        if let Some(Value::Object(children_arr)) = n.properties.get("children") {
            let children_arr = children_arr.clone();
            if let ObjectKind::Array(ref mut items) = children_arr.lock().unwrap().kind {
                *items = element_children;
            }
        } else {
            n.properties.insert("children".into(), make_array(element_children));
        }
        n.properties.insert("firstElementChild".into(), first_elem);
        n.properties.insert("lastElementChild".into(), last_elem);
    } else if is_text_like {
        let nv = n.properties.get("nodeValue").cloned().unwrap_or(Value::Null);
        n.properties.insert("textContent".into(), nv);
    }
    // Drop the lock BEFORE recursing into `collect_text`, which
    // re-acquires the same node's mutex. std Mutex isn't reentrant
    // — holding the lock here would deadlock the textContent walk.
    drop(n);

    if is_element_or_doc {
        let mut buf = String::new();
        collect_text(node, &mut buf);
        node.lock().unwrap().properties.insert("textContent".into(), s(&buf));
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
    let Some(Value::Object(arr)) = element else { return };
    let elem_root = {
        let a = arr.lock().unwrap();
        if let ObjectKind::Array(ref items) = a.kind {
            items.iter().find(|v| matches!(v, Value::Object(o)
                if matches!(o.lock().unwrap().properties.get("nodeType"), Some(Value::I32(1)))))
                .cloned()
                .unwrap_or(Value::Null)
        } else {
            Value::Null
        }
    };
    doc.lock().unwrap().properties.insert("documentElement".into(), elem_root);
}

// ── Tree walkers (computed methods) ───────────────────────────────

fn find_by_id(node: &Arc<Mutex<Object>>, id: &str) -> Option<Value> {
    let n = node.lock().unwrap();
    if let Some(Value::String(node_id)) = n.properties.get("id") {
        if node_id.as_ref() == id {
            return Some(Value::Object(node.clone()));
        }
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
        } else { return };
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
        } else { return };
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
                } else { return };
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
                    if k.starts_with("__") { continue; }
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
                    } else { vec![] }
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
    Value::Object(Arc::new(Mutex::new(Object::new_array(elems))))
}

fn empty_array() -> Value {
    make_array(vec![])
}

fn make_empty_object() -> Value {
    let mut o = Object::new();
    o.properties.insert("__type".into(), s("NamedNodeMap"));
    Value::Object(Arc::new(Mutex::new(o)))
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
            if start == i { return None; }
            parts.push(SimplePart::Id(chars[start..i].iter().collect()));
        } else if c == '.' {
            i += 1;
            let start = i;
            while i < chars.len() && is_ident_char(chars[i]) {
                i += 1;
            }
            if start == i { return None; }
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
    if start == i { return None; }
    let name: String = chars[start..i].iter().collect();
    // Skip whitespace
    while i < chars.len() && chars[i].is_whitespace() { i += 1; }
    if i >= chars.len() { return None; }
    if chars[i] == ']' {
        return Some((SimplePart::Attr { name, op: AttrOp::Has, value: None }, i + 1));
    }
    // Operator: =, *=, ^=, $=, ~=, |=
    let op = match chars[i] {
        '=' => { i += 1; AttrOp::Exact }
        '*' if i + 1 < chars.len() && chars[i + 1] == '=' => { i += 2; AttrOp::Substring }
        '^' if i + 1 < chars.len() && chars[i + 1] == '=' => { i += 2; AttrOp::Prefix }
        '$' if i + 1 < chars.len() && chars[i + 1] == '=' => { i += 2; AttrOp::Suffix }
        '~' if i + 1 < chars.len() && chars[i + 1] == '=' => { i += 2; AttrOp::Word }
        '|' if i + 1 < chars.len() && chars[i + 1] == '=' => { i += 2; AttrOp::Lang }
        _ => return None,
    };
    // Skip whitespace
    while i < chars.len() && chars[i].is_whitespace() { i += 1; }
    if i >= chars.len() { return None; }
    // Quoted or unquoted value
    let value = if chars[i] == '"' || chars[i] == '\'' {
        let quote = chars[i];
        i += 1;
        let start = i;
        while i < chars.len() && chars[i] != quote {
            i += 1;
        }
        if i >= chars.len() { return None; }
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
    while i < chars.len() && chars[i].is_whitespace() { i += 1; }
    if i >= chars.len() || chars[i] != ']' { return None; }
    Some((SimplePart::Attr { name, op, value: Some(value) }, i + 1))
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
    if parts.is_empty() { return true; }
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
                    if let Some(pp) = node.clone() { current = pp.clone(); }
                    else { break; }
                    node = parent_of(&current);
                }
                if !found { return false; }
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
                if !found { return false; }
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

fn find_first_match(
    root: &Arc<Mutex<Object>>,
    selectors: &[ComplexSelector],
) -> Option<Value> {
    if selectors.iter().any(|s| selector_matches(root, s))
        && {
            let n = root.lock().unwrap();
            matches!(n.properties.get("nodeType"), Some(Value::I32(t)) if *t == ELEMENT_NODE)
        }
    {
        return Some(Value::Object(root.clone()));
    }
    let children = {
        let n = root.lock().unwrap();
        n.properties.get("childNodes").cloned()
    };
    let Some(Value::Object(arr)) = children else { return None };
    let items: Vec<Value> = if let ObjectKind::Array(ref a) = arr.lock().unwrap().kind {
        a.clone()
    } else { return None };
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
    let Some(Value::Object(arr)) = children else { return };
    let items: Vec<Value> = if let ObjectKind::Array(ref a) = arr.lock().unwrap().kind {
        a.clone()
    } else { return };
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
    o.properties.insert("nodeType".into(), Value::I32(ELEMENT_NODE));
    o.properties.insert("nodeName".into(), s(tag));
    o.properties.insert("tagName".into(), s(tag));
    o.properties.insert("localName".into(), s(&local_name));
    o.properties.insert("prefix".into(), prefix.map(|p| s(&p)).unwrap_or(Value::Null));
    o.properties.insert("namespaceURI".into(), Value::Null);
    o.properties.insert("nodeValue".into(), Value::Null);

    let mut attrs = Object::new();
    attrs.type_id = dom_type_ids().map(|d| d.named_node_map).unwrap_or(0);
    attrs.properties.insert("__type".into(), s("NamedNodeMap"));
    o.properties.insert("attributes".into(), Value::Object(Arc::new(Mutex::new(attrs))));
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
    o.properties.insert("ownerDocument".into(),
        owner.map(|d| Value::Object(d.clone())).unwrap_or(Value::Null));
    Value::Object(Arc::new(Mutex::new(o)))
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
    let Some(Value::Object(arr)) = arr else { return };
    {
        let mut a = arr.lock().unwrap();
        if let ObjectKind::Array(ref mut items) = a.kind {
            items.push(Value::Object(child.clone()));
        }
    }
    // Update child's parent.
    {
        let mut c = child.lock().unwrap();
        c.properties.insert("parentNode".into(), Value::Object(parent.clone()));
    }
    refresh_node_relationships(parent);
}

fn remove_child_inner(parent: &Arc<Mutex<Object>>, child: &Arc<Mutex<Object>>) -> bool {
    let arr = {
        let p = parent.lock().unwrap();
        p.properties.get("childNodes").cloned()
    };
    let Some(Value::Object(arr)) = arr else { return false };
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
    let Some(Value::Object(arr)) = arr else { return false };
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
        new_child.lock().unwrap().properties.insert("parentNode".into(), Value::Object(parent.clone()));
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
    let Some(Value::Object(arr)) = children else { return };
    let items: Vec<Value> = if let ObjectKind::Array(ref a) = arr.lock().unwrap().kind {
        a.clone()
    } else {
        return;
    };

    for (i, child) in items.iter().enumerate() {
        if let Value::Object(child_obj) = child {
            let prev = if i > 0 { items.get(i - 1).cloned() } else { None };
            let next = items.get(i + 1).cloned();
            let mut c = child_obj.lock().unwrap();
            c.properties.insert("previousSibling".into(), prev.unwrap_or(Value::Null));
            c.properties.insert("nextSibling".into(), next.unwrap_or(Value::Null));
        }
    }

    let element_children: Vec<Value> = items.iter()
        .filter(|v| matches!(v, Value::Object(o)
            if matches!(o.lock().unwrap().properties.get("nodeType"), Some(Value::I32(1)))))
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
                if let ObjectKind::Array(ref mut a_items) = children_arr.lock().unwrap().kind {
                    *a_items = element_children;
                }
            } else {
                p.properties.insert("children".into(), make_array(element_children));
            }
            p.properties.insert("firstElementChild".into(), first_elem);
            p.properties.insert("lastElementChild".into(), last_elem);
            // documentElement on Document always tracks first Element child.
            if matches!(p.properties.get("nodeType"), Some(Value::I32(t)) if *t == DOCUMENT_NODE) {
                let elem_root = items.iter().find(|v| matches!(v, Value::Object(o)
                    if matches!(o.lock().unwrap().properties.get("nodeType"), Some(Value::I32(1)))))
                    .cloned()
                    .unwrap_or(Value::Null);
                p.properties.insert("documentElement".into(), elem_root);
            }
        }
    }
    // textContent on the parent — needs lock dropped before recursing.
    let mut buf = String::new();
    collect_text(parent, &mut buf);
    parent.lock().unwrap().properties.insert("textContent".into(), s(&buf));
}

fn clone_node(node: &Arc<Mutex<Object>>, deep: bool) -> Value {
    let n = node.lock().unwrap();
    let mut clone = Object::new();
    clone.type_id = n.type_id;
    // Shallow copy of all properties EXCEPT structural pointers.
    for (key, val) in &n.properties {
        match key.as_str() {
            "parentNode" | "nextSibling" | "previousSibling"
            | "firstChild" | "lastChild" | "firstElementChild" | "lastElementChild"
            | "ownerDocument" => continue, // reset structural state
            "childNodes" | "children" | "attributes" => continue, // handled below
            _ => {
                clone.properties.insert(key.clone(), val.clone());
            }
        }
    }
    // Reset structural state.
    clone.properties.insert("parentNode".into(), Value::Null);
    clone.properties.insert("nextSibling".into(), Value::Null);
    clone.properties.insert("previousSibling".into(), Value::Null);
    clone.properties.insert("firstChild".into(), Value::Null);
    clone.properties.insert("lastChild".into(), Value::Null);
    clone.properties.insert("firstElementChild".into(), Value::Null);
    clone.properties.insert("lastElementChild".into(), Value::Null);
    clone.properties.insert("ownerDocument".into(), Value::Null);

    // Clone attributes (deep copy of NamedNodeMap).
    if let Some(Value::Object(attrs_src)) = n.properties.get("attributes") {
        let mut attrs_clone = Object::new();
        attrs_clone.type_id = attrs_src.lock().unwrap().type_id;
        for (k, v) in &attrs_src.lock().unwrap().properties {
            attrs_clone.properties.insert(k.clone(), v.clone());
        }
        clone.properties.insert("attributes".into(),
            Value::Object(Arc::new(Mutex::new(attrs_clone))));
    }
    clone.properties.insert("childNodes".into(), make_array(vec![]));
    clone.properties.insert("children".into(), make_array(vec![]));
    drop(n);

    let clone_arc = Arc::new(Mutex::new(clone));
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
    o.properties.insert("nodeType".into(), Value::I32(DOCUMENT_FRAGMENT_NODE));
    o.properties.insert("nodeName".into(), s("#document-fragment"));
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
    o.properties.insert("ownerDocument".into(),
        owner.map(|d| Value::Object(d.clone())).unwrap_or(Value::Null));
    Value::Object(Arc::new(Mutex::new(o)))
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
        } else { return };
        for child in items {
            if let Value::Object(c) = child {
                find_by_local_name(&c, local, out);
            }
        }
    }
}
