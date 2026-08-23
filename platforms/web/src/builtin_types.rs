//! WHATWG / W3C web-platform **types** — the runtime TypeRegistry vtables for
//! `TextEncoder`/`TextDecoder`, `URLSearchParams`, fetch `Response`, and the
//! DOM node hierarchy (`Element`/`Document`/`Text`/…).
//!
//! The `register_type` counterpart to the web plugin's host-fn `init`: the web
//! plugin declares its own types here, in its `finalize`. Each method resolves
//! a `web:*` host fn by registry index, so it runs after every plugin's `init`.

use vybe_runtime::Framework;
use vybe_runtime::{Method, TypeDef};

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
    // `XMLDocument`, not `Document` — the spec's own name for what
    // `parseFromString` answers for the XML content types (DOM §4.5.1), and it
    // has to be a distinct type now that `Document` is the LIVE document
    // registered below. One `Document` TypeDef cannot serve both: pointed at
    // `web:dom`, a property-bag document would still carry that type id,
    // `doc.getElementById(…)` would dispatch into the rendering engine, find no
    // `__document` on the bag, and answer about document 0 — the wrong answer,
    // silently.
    //
    // The bag stays for the XML path on purpose: PHP's `DOMDocument`, .NET
    // XLinq and `xml2js` all read `nodeType`/`childNodes`/`attributes` straight
    // off the object, and namespaces, `Attr` nodes, CDATA and PIs exist only
    // there. `parseFromString(s, "text/html")` already answers a live handle.
    let xml_document_id = {
        let mut t = TypeDef::new("XMLDocument");
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

    // ── The LIVE document: HTMLDocument / HTMLElement ───────────────
    //
    // The types above belong to `web:dom-parser` — detached `Value::Object`
    // trees, which is the right shape for `DOMParser().parseFromString(…)` and
    // `XMLHttpRequest.responseXML`, and which render nothing. The document a
    // page actually has is `vybe_widgets::dom`, reached through `web:dom` /
    // `web:html` / `web:cssom`, and it needs its own vtables because the same
    // method name has a different implementation over a different tree.
    //
    // Naming is the spec's, in both directions: the tree a page has is the
    // `Document`, and its elements are `HTMLElement`s; the parsed XML tree
    // above is an `XMLDocument` whose elements are plain `Element`s.
    //
    // `parent` is `Object` rather than `Element`, deliberately — inheriting
    // there would make any method not listed below resolve into the OTHER tree
    // and quietly answer about the wrong document, which is worse than not
    // resolving at all.
    //
    // `insertBefore` / `replaceChild` / `cloneNode` are absent on purpose:
    // `vybe_widgets::dom::Document` has `append_child` and `remove_child` and
    // nothing else, so there is no engine operation to forward to. They are not
    // stubbed — a missing method fails to resolve, which is visible.
    let html_element_id = {
        let mut t = TypeDef::new("HTMLElement");
        // (method, module, host fn). The module differs per method because the
        // SPEC splits them: `appendChild` is DOM, `focus`/`value` are the HTML
        // element IDL, `style.setProperty` is CSSOM.
        for (method, module, fname) in &[
            // DOM core
            ("appendChild", "web:dom", "appendChild"),
            ("removeChild", "web:dom", "removeChild"),
            ("setTextContent", "web:dom", "setTextContent"),
            ("textContent", "web:dom", "textContent"),
            ("isConnected", "web:dom", "isConnected"),
            ("setAttribute", "web:dom", "setAttribute"),
            ("getAttribute", "web:dom", "getAttribute"),
            // ⚠ This was mapped to `toggleAttribute` — asking whether an
            // attribute exists FLIPPED it, and answered the toggle's result
            // rather than the presence. `hasAttribute` is a real host fn now.
            ("hasAttribute", "web:dom", "hasAttribute"),
            // `getAttributeNames()` — DOM §4.9. The only attribute call that is
            // not addressed by name, which is what lets a caller ask an element
            // what it HAS rather than whether it has one particular thing.
            ("getAttributeNames", "web:dom", "getAttributeNames"),
            ("toggleAttribute", "web:dom", "toggleAttribute"),
            ("removeAttribute", "web:dom", "removeAttribute"),
            // Traversal — DOM §4.2.6 / §4.4, over the LIVE tree.
            ("firstChild", "web:dom", "firstChild"),
            ("lastChild", "web:dom", "lastChild"),
            ("nextSibling", "web:dom", "nextSibling"),
            ("previousSibling", "web:dom", "previousSibling"),
            ("parentNode", "web:dom", "parentNode"),
            ("childNodes", "web:dom", "childNodes"),
            ("children", "web:dom", "children"),
            ("firstElementChild", "web:dom", "firstElementChild"),
            ("lastElementChild", "web:dom", "lastElementChild"),
            // Mutation and matching.
            ("remove", "web:dom", "remove"),
            ("matches", "web:dom", "matches"),
            ("closest", "web:dom", "closest"),
            ("contains", "web:dom", "contains"),
            ("querySelector", "web:dom", "querySelector"),
            ("querySelectorAll", "web:dom", "querySelectorAll"),
            // `classList` — flat, because a host vtable has nowhere to hang a
            // live `DOMTokenList`. `el.classListAdd("x")` reaches the same
            // store `class` does; the JS spelling `el.classList.add("x")`
            // needs a `DOMTokenList` object and is NOT wired.
            ("classListAdd", "web:dom", "classListAdd"),
            ("classListRemove", "web:dom", "classListRemove"),
            ("classListToggle", "web:dom", "classListToggle"),
            ("classListContains", "web:dom", "classListContains"),
            ("removeAttribute", "web:dom", "removeAttribute"),
            ("addEventListener", "web:dom", "addEventListener"),
            ("removeEventListener", "web:dom", "removeEventListener"),
            ("innerHtml", "web:dom", "innerHtml"),
            ("setInnerHtml", "web:dom", "setInnerHtml"),
            ("querySelector", "web:dom", "querySelector"),
            ("querySelectorAll", "web:dom", "querySelectorAll"),
            ("getElementsByTagName", "web:dom", "getElementsByTagName"),
            // Traversal and matching — DOM §4.2.6 / §4.9. Registered on the
            // ELEMENT vtable so a guest reaches them by their IDL spelling,
            // which is the whole point of the 1:1 mapping: `el.closest(sel)`
            // in the guest IS `closest` in the host, not a shim.
            ("firstChild", "web:dom", "firstChild"),
            ("lastChild", "web:dom", "lastChild"),
            ("nextSibling", "web:dom", "nextSibling"),
            ("previousSibling", "web:dom", "previousSibling"),
            ("children", "web:dom", "children"),
            ("firstElementChild", "web:dom", "firstElementChild"),
            ("lastElementChild", "web:dom", "lastElementChild"),
            ("remove", "web:dom", "remove"),
            ("hasAttribute", "web:dom", "hasAttribute"),
            // `getAttributeNames()` — DOM §4.9. The only attribute call that is
            // not addressed by name, which is what lets a caller ask an element
            // what it HAS rather than whether it has one particular thing.
            ("getAttributeNames", "web:dom", "getAttributeNames"),
            ("matches", "web:dom", "matches"),
            ("closest", "web:dom", "closest"),
            ("contains", "web:dom", "contains"),
            // HTML element IDL
            //
            // `canvas.getContext("2d")` — HTML §4.12.5. It belongs on the
            // ELEMENT because that is where the IDL puts it, and a page has no
            // other way to reach a canvas: `getElementById(…).getContext("2d")`
            // is THE spelling, and without this entry it answered undefined
            // while the host function sat there registered and reachable only
            // as a free call no page would ever write.
            ("getContext", "web:canvas", "getContext"),
            ("focus", "web:html", "focus"),
            ("showPicker", "web:html", "showPicker"),
            ("show", "web:html", "show"),
            ("showModal", "web:html", "showModal"),
            ("close", "web:html", "close"),
            // CSSOM
            ("setStyleProperty", "web:cssom", "setStyleProperty"),
            ("getStyleProperty", "web:cssom", "getStyleProperty"),
            // The resolved-value read. A frontend that wants the laid-out pixel
            // (a toolkit's `Left`/`Top`/`Width`/`Height`) names THIS one;
            // `getStyleProperty` above answers the declared value.
            (
                "getComputedStyleProperty",
                "web:cssom",
                "getComputedStyleProperty",
            ),
            ("removeStyleProperty", "web:cssom", "removeStyleProperty"),
            // ⚠ `classList` is NOT here, and cannot be a flat vtable entry:
            // in the IDL it is a PROPERTY returning a live `DOMTokenList`, so
            // `el.classList.add(x)` is two hops — a property read, then a
            // method call on what it answers. The host functions exist
            // (`classListAdd`/`Remove`/`Toggle`/`Contains`) and a frontend can
            // route a `classList` role straight to them; what is missing is
            // the `DOMTokenList` object for the JS spelling. Mapping
            // `classList` to a method here would make `el.classList` callable
            // and `el.classList.add` undefined, which is worse than absent.
        ] {
            if let Some(idx) = fw.host_fn_index(module, fname) {
                t.methods.insert(method.to_string(), Method::HostFn(idx));
            }
        }
        t.parent = Some(0);
        fw.register_type(t)
    };
    let html_document_id = {
        let mut t = TypeDef::new("Document");
        for (method, module, fname) in &[
            ("createElement", "web:dom", "createElement"),
            ("createTextNode", "web:dom", "createTextNode"),
            ("createComment", "web:dom", "createComment"),
            ("getElementById", "web:dom", "getElementById"),
            ("querySelector", "web:dom", "querySelector"),
            ("querySelectorAll", "web:dom", "querySelectorAll"),
            ("getElementsByTagName", "web:dom", "getElementsByTagName"),
            ("appendChild", "web:dom", "appendChild"),
            ("removeChild", "web:dom", "removeChild"),
            ("addEventListener", "web:dom", "addEventListener"),
            ("removeEventListener", "web:dom", "removeEventListener"),
            ("innerHtml", "web:dom", "innerHtml"),
            ("setInnerHtml", "web:dom", "setInnerHtml"),
            ("setTextContent", "web:dom", "setTextContent"),
            ("title", "web:html", "title"),
            ("setTitle", "web:html", "setTitle"),
        ] {
            if let Some(idx) = fw.host_fn_index(module, fname) {
                t.methods.insert(method.to_string(), Method::HostFn(idx));
            }
        }
        t.parent = Some(0);
        fw.register_type(t)
    };
    // ── CanvasRenderingContext2D ───────────────────────────────────
    //
    // What `getContext("2d")` answers. The object was already built with
    // `__type: "CanvasRenderingContext2D"` stamped on it (`canvas.rs`), and
    // `Vm::get_property` resolves a vtable BY that `__type` — so the string was
    // naming a type that did not exist and every `ctx.fillRect(…)` answered
    // undefined. A page could reach `web:canvas` only by calling the host
    // functions flat, which is a spelling no page has.
    //
    // Each entry is the IDL's own name mapped to the host function that was
    // already registered for it: the vtable IS the interface, so a member
    // absent here is one the seam cannot carry yet rather than one spelled
    // differently.
    //
    // ⚠ The `setXxx` entries are NOT IDL members. `fillStyle`, `strokeStyle`,
    // `font`, `lineWidth` and the rest are ATTRIBUTES, and an attribute
    // assignment is `__set_<name>`, not a method. They cannot become accessors
    // until the seam carries the CSS STRING a page assigns — `Op2D` holds
    // `SetFillStyle(u8, u8, u8, u8)`, a colour already parsed, and
    // `platforms/web` deliberately parses no CSS (`web:cssom` forwards
    // property text to the engine for exactly this reason). Until then these
    // are how a caller sets state, and they are the reason this type is usable
    // rather than able to draw in black only.
    {
        let mut t = TypeDef::new("CanvasRenderingContext2D");
        for (member, fname) in &[
            // §4.12.5.1.2 state
            ("save", "save"),
            ("restore", "restore"),
            // §4.12.5.1.3 transformations
            ("scale", "scale"),
            ("rotate", "rotate"),
            ("translate", "translate"),
            ("transform", "transform"),
            ("setTransform", "setTransform"),
            ("resetTransform", "resetTransform"),
            // §4.12.5.1.8 building paths
            ("beginPath", "beginPath"),
            ("closePath", "closePath"),
            ("moveTo", "moveTo"),
            ("lineTo", "lineTo"),
            ("quadraticCurveTo", "quadraticCurveTo"),
            ("bezierCurveTo", "bezierCurveTo"),
            ("arc", "arc"),
            ("ellipse", "ellipse"),
            ("rect", "rect"),
            // §4.12.5.1.9 / §4.12.5.1.5 drawing
            ("fill", "fill"),
            ("stroke", "stroke"),
            ("clip", "clip"),
            ("clearRect", "clearRect"),
            ("fillRect", "fillRect"),
            ("strokeRect", "strokeRect"),
            // §4.12.5.1.11 text
            ("fillText", "fillText"),
            ("strokeText", "strokeText"),
            ("measureText", "measureText"),
            // §4.12.5.1.12 images and pixel manipulation
            ("drawImage", "drawImage"),
            ("createImageData", "createImageData"),
            ("putImageData", "putImageData"),
            // §4.12.5.1.6 line styles
            ("setLineDash", "setLineDash"),
            // §4.12.5.1.7 gradients and patterns — the factories. A page
            // cannot build a gradient any other way, so without these the
            // seam could carry one and nothing could make one.
            ("createLinearGradient", "createLinearGradient"),
            ("createRadialGradient", "createRadialGradient"),
            ("createConicGradient", "createConicGradient"),
            ("createPattern", "createPattern"),
            // §4.12.5.1.9 — `new Path2D()` reached as a factory, because this
            // platform has no `new` for a host type yet. `addPath` on the
            // CONTEXT folds a prebuilt path into the current one.
            ("createPath2D", "createPath2D"),
            ("addPath", "appendPath"),
            // §4.12.5.1.13 shadows
            ("setShadowBlur", "setShadowBlur"),
            ("setShadowOffsetX", "setShadowOffsetX"),
            ("setShadowOffsetY", "setShadowOffsetY"),
            ("setShadowColor", "setShadowColor"),
            // §4.12.5.1.8 the paths the seam could not express
            ("arcTo", "arcTo"),
            ("roundRect", "roundRect"),
            ("ellipseFull", "ellipseFull"),
            ("fillWithRule", "fillWithRule"),
            ("clipWithRule", "clipWithRule"),
            ("reset", "reset"),
            ("drawFocusIfNeeded", "drawFocusIfNeeded"),
            ("fillTextMaxWidth", "fillTextMaxWidth"),
            ("strokeTextMaxWidth", "strokeTextMaxWidth"),
            // The half that ASKS — none of these were reachable at all.
            ("isPointInPath", "isPointInPath"),
            ("isPointInStroke", "isPointInStroke"),
            ("isContextLost", "isContextLost"),
            ("getLineDash", "getLineDash"),
            ("getTransform", "getTransform"),
            ("getImageData", "getImageData"),
            ("toDataURL", "toDataURL"),
            ("toBlob", "toBlob"),
            ("getContextAttributes", "getContextAttributes"),
            // Attribute read-back. §4.12.5 requires every one of these to
            // serialize; a page could set them and never ask.
            ("getFont", "getFont"),
            ("getFillStyle", "getFillStyle"),
            ("getStrokeStyle", "getStrokeStyle"),
            ("getFilter", "getFilter"),
            ("getGlobalCompositeOperation", "getGlobalCompositeOperation"),
            ("getImageSmoothingQuality", "getImageSmoothingQuality"),
            ("getShadowColor", "getShadowColor"),
            ("getDirection", "getDirection"),
            ("getLetterSpacing", "getLetterSpacing"),
            ("getWordSpacing", "getWordSpacing"),
            ("getFontKerning", "getFontKerning"),
            ("getFontStretch", "getFontStretch"),
            ("getFontVariantCaps", "getFontVariantCaps"),
            ("getTextRendering", "getTextRendering"),
            ("getLang", "getLang"),
            ("getTextAlign", "getTextAlign"),
            ("getTextBaseline", "getTextBaseline"),
            ("getLineCap", "getLineCap"),
            ("getLineJoin", "getLineJoin"),
            // Attribute writes that take a CSS STRING, which is what the IDL
            // says a page assigns. The engine parses them.
            ("setFillStyleCss", "setFillStyleCss"),
            ("setStrokeStyleCss", "setStrokeStyleCss"),
            ("setFontCss", "setFontCss"),
            ("setFilter", "setFilter"),
            ("setGlobalCompositeOperation", "setGlobalCompositeOperation"),
            ("setImageSmoothingQuality", "setImageSmoothingQuality"),
            ("setDirection", "setDirection"),
            ("setLetterSpacing", "setLetterSpacing"),
            ("setWordSpacing", "setWordSpacing"),
            ("setFontKerning", "setFontKerning"),
            ("setFontStretch", "setFontStretch"),
            ("setFontVariantCaps", "setFontVariantCaps"),
            ("setTextRendering", "setTextRendering"),
            ("setLang", "setLang"),
            // Pre-attribute spellings — see the note above.
            ("setFillStyle", "setFillStyle"),
            ("setStrokeStyle", "setStrokeStyle"),
            ("setLineWidth", "setLineWidth"),
            ("setLineCap", "setLineCap"),
            ("setLineJoin", "setLineJoin"),
            ("setMiterLimit", "setMiterLimit"),
            ("setLineDashOffset", "setLineDashOffset"),
            ("setGlobalAlpha", "setGlobalAlpha"),
            ("setFont", "setFont"),
            ("setTextAlign", "setTextAlign"),
            ("setTextBaseline", "setTextBaseline"),
            ("setImageSmoothingEnabled", "setImageSmoothingEnabled"),
        ] {
            if let Some(idx) = fw.host_fn_index("web:canvas", fname) {
                t.methods.insert(member.to_string(), Method::HostFn(idx));
            }
        }
        t.parent = Some(0);
        fw.register_type(t);
    }

    // ── Path2D ─────────────────────────────────────────────────────
    //
    // The same path-building names a context has, on a different receiver.
    // That is the interface's whole point: a path is described identically
    // whether it is going into the context's current path or into a shape
    // being kept for later, so a page does not learn two vocabularies.
    {
        let mut t = TypeDef::new("Path2D");
        for (member, fname) in &[
            ("closePath", "pathClosePath"),
            ("moveTo", "pathMoveTo"),
            ("lineTo", "pathLineTo"),
            ("quadraticCurveTo", "pathQuadraticCurveTo"),
            ("bezierCurveTo", "pathBezierCurveTo"),
            ("arcTo", "pathArcTo"),
            ("rect", "pathRect"),
            ("roundRect", "pathRoundRect"),
            ("arc", "pathArc"),
            ("ellipse", "pathEllipse"),
            ("addPath", "addPath"),
        ] {
            if let Some(idx) = fw.host_fn_index("web:canvas", fname) {
                t.methods.insert(member.to_string(), Method::HostFn(idx));
            }
        }
        t.parent = Some(0);
        fw.register_type(t);
    }

    // ── CanvasGradient ─────────────────────────────────────────────
    //
    // One method, and the reason the gradient factories can answer a plain
    // object: `addColorStop` mutates what the page is holding, and the whole
    // definition crosses the seam later when it is assigned to `fillStyle`.
    {
        let mut t = TypeDef::new("CanvasGradient");
        if let Some(idx) = fw.host_fn_index("web:canvas", "addColorStop") {
            t.methods.insert("addColorStop".to_string(), Method::HostFn(idx));
        }
        t.parent = Some(0);
        fw.register_type(t);
    }
    // `CanvasPattern` has one method in the IDL, `setTransform`, which this
    // seam cannot carry yet — registered so the type EXISTS and a page can
    // hold one and assign it, which is what a pattern is for.
    {
        let mut t = TypeDef::new("CanvasPattern");
        t.parent = Some(0);
        fw.register_type(t);
    }

    crate::html::set_live_type_ids(crate::html::LiveTypeIds {
        document: html_document_id,
        element: html_element_id,
    });

    // Hand the IDs to `dom_parser` so the parser stamps each constructed node's
    // `Object::type_id` for vtable dispatch.
    crate::dom_parser::set_dom_type_ids(crate::dom_parser::DomTypeIds {
        document: xml_document_id,
        element: element_id,
        text: text_id,
        cdata: cdata_id,
        comment: comment_id,
        processing_instruction: pi_id,
        attr: attr_id,
        named_node_map: nnm_id,
    });
}
