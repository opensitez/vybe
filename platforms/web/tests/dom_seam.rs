//! The `web:html` seam end to end: operations in, spec-shaped answers out,
//! and a real user click coming back as a `click` on the right node.
//!
//! Drives [`DomOp`] directly rather than through the VM, so what is under
//! test is the surface's contract with the engine — the part a browser
//! backend would have to satisfy too.

use vybe_platform_web::engine::{DOCUMENT, DomOp, DomValue, apply, new_document};

fn node(v: DomValue) -> u64 {
    match v {
        DomValue::Node(n) => n,
        other => panic!("expected a node, got {:?}", other),
    }
}

fn text(v: DomValue) -> String {
    match v {
        DomValue::Text(s) => s,
        other => panic!("expected text, got {:?}", other),
    }
}

fn create(doc: u64, tag: &str, input_type: &str) -> u64 {
    node(apply(
        doc,
        DomOp::CreateElement {
            tag: tag.into(),
            input_type: input_type.into(),
        },
    ))
}

/// Install whichever browser this build selected.
///
/// The ONE engine-specific line in this file. Everything below is the WHATWG
/// contract, so the same assertions run against either browser and there is no
/// second copy to drift:
///
///     cargo test -p vybe_platform_web --features gui             # widgets
///     cargo test -p vybe_platform_web --features engine-webcore  # webcore
fn install() {
    #[cfg(feature = "engine-webcore")]
    vybe_platform_web::engine_webcore::install();
    #[cfg(not(feature = "engine-webcore"))]
    vybe_platform_web::engine_widgets::install();
}

fn setup() -> u64 {
    install();
    new_document("test")
}

#[test]
fn a_created_element_is_not_in_the_document() {
    let doc = setup();
    let cb = create(doc, "input", "checkbox");
    assert!(
        matches!(apply(doc, DomOp::IsConnected(cb)), DomValue::Bool(false)),
        "createElement must not insert"
    );
    apply(
        doc,
        DomOp::AppendChild {
            parent: DOCUMENT,
            child: cb,
        },
    );
    assert!(matches!(
        apply(doc, DomOp::IsConnected(cb)),
        DomValue::Bool(true)
    ));
}

#[test]
fn an_absent_attribute_is_null_not_empty_string() {
    let doc = setup();
    let b = create(doc, "button", "");
    assert!(
        matches!(
            apply(doc, DomOp::GetAttribute(b, "class".into())),
            DomValue::Null
        ),
        "absent attribute must be null"
    );
    apply(
        doc,
        DomOp::SetAttribute(b, "class".into(), "primary".into()),
    );
    assert_eq!(
        text(apply(doc, DomOp::GetAttribute(b, "class".into()))),
        "primary"
    );
}

#[test]
fn checked_is_a_boolean_and_value_a_string() {
    let doc = setup();
    let cb = create(doc, "input", "checkbox");
    apply(
        doc,
        DomOp::AppendChild {
            parent: DOCUMENT,
            child: cb,
        },
    );
    assert!(matches!(
        apply(doc, DomOp::Checked(cb)),
        DomValue::Bool(false)
    ));
    apply(doc, DomOp::SetChecked(cb, true));
    assert!(matches!(
        apply(doc, DomOp::Checked(cb)),
        DomValue::Bool(true)
    ));

    let field = create(doc, "input", "text");
    apply(doc, DomOp::SetValue(field, "42".into()));
    // A string, not a number and not `"True"`-style stringification.
    assert_eq!(text(apply(doc, DomOp::Value(field))), "42");
}

#[test]
fn a_range_input_reports_its_number_as_value() {
    let doc = setup();
    let r = create(doc, "input", "range");
    apply(
        doc,
        DomOp::AppendChild {
            parent: DOCUMENT,
            child: r,
        },
    );
    apply(doc, DomOp::SetValue(r, "30".into()));
    // An integer `value` carries no `.0` — IDL value is a DOMString.
    assert_eq!(text(apply(doc, DomOp::Value(r))), "30");
}

#[test]
fn select_options_are_the_elements_content() {
    let doc = setup();
    let s = create(doc, "select", "");
    apply(
        doc,
        DomOp::AppendChild {
            parent: DOCUMENT,
            child: s,
        },
    );
    apply(doc, DomOp::AddItem(s, "one".into()));
    apply(doc, DomOp::AddItem(s, "two".into()));

    // The items ARE the element's content — `select.options[i].text`.
    assert_eq!(text(apply(doc, DomOp::ItemText(s, 0))), "one");
    assert_eq!(text(apply(doc, DomOp::ItemText(s, 1))), "two");

    // ⚠ This asserted `SetValue("1")` then `Value == "1"` — which passed
    // whether or not `AddItem` did anything, because it only round-tripped a
    // string. `select.value = v` selects the option WORTH v (HTML §4.10.7);
    // the index is `selectedIndex`, its own IDL member.
    apply(doc, DomOp::SetValue(s, "two".into()));
    assert_eq!(text(apply(doc, DomOp::Value(s))), "two");
    assert!(
        matches!(apply(doc, DomOp::SelectedIndex(s)), DomValue::Number(n) if n == 1.0),
        "assigning a value must move selectedIndex to that option"
    );
}

#[test]
fn style_uses_css_units() {
    let doc = setup();
    let b = create(doc, "button", "");
    apply(
        doc,
        DomOp::AppendChild {
            parent: DOCUMENT,
            child: b,
        },
    );
    apply(
        doc,
        DomOp::SetStyleProperty(b, "left".into(), "40px".into()),
    );
    apply(doc, DomOp::SetStyleProperty(b, "top".into(), "1em".into()));
    assert_eq!(
        text(apply(doc, DomOp::GetStyleProperty(b, "left".into()))),
        "40px"
    );
    // CSSOM §6.4.2: `element.style.getPropertyValue()` serializes the DECLARED
    // value. `1em` was authored, so `1em` reads back.
    //
    // This asserted `"16px"` until the resolved read got its own op. The engine
    // was answering `left`/`top`/`width`/`height` out of the laid-out rect —
    // `getComputedStyle`'s job — so a stylesheet round-trip through
    // `getPropertyValue` silently rewrote authored units. webcore always
    // answered the declared value, which is how the divergence surfaced.
    assert_eq!(
        text(apply(doc, DomOp::GetStyleProperty(b, "top".into()))),
        "1em",
        "getPropertyValue must not resolve against layout"
    );
    // The resolved answer still exists — under the name that means it.
    assert_eq!(
        text(apply(doc, DomOp::ComputedStyleProperty(b, "top".into()))),
        "16px"
    );
}

#[test]
fn two_documents_are_two_trees() {
    install();
    let a = new_document("a");
    let b = new_document("b");
    assert_ne!(a, b);
    let n = create(a, "button", "");
    apply(a, DomOp::SetAttribute(n, "id".into(), "shared".into()));
    assert!(
        matches!(
            apply(b, DomOp::GetElementById("shared".into())),
            DomValue::Null
        ),
        "an element in one document must not be findable in another"
    );
}

/// Synthesise a real user click on `node`, however THIS browser takes input.
///
/// Click the middle of a node, as the window shell does.
///
/// **No longer engine-specific.** This was two functions behind a `#[cfg]` —
/// one poking `widgets`' form, one calling webcore's `process_mouse_event`
/// — because delivering OS input was the one thing the seam had no verb for.
/// The host had the same hole and filled it the same way, by reaching past the
/// engine into the toolkit, so a click under any other engine went nowhere.
/// `DispatchPointer` is that verb, and this is now one path for both.
fn click_at(doc: u64, node: u64) {
    // Hit-testing needs geometry, and a document that has never been laid out
    // has none. Asking for a resolved value is what forces layout in a browser
    // too — `getComputedStyle` is a documented reflow trigger — so this is the
    // page-level way to say "settle first" rather than an engine call.
    apply(doc, DomOp::ComputedStyleProperty(node, "width".into()));
    // The element's CENTRE, so the click lands wherever layout put it rather
    // than at a coordinate the test guessed.
    let (x, y) = match apply(doc, DomOp::BoundingClientRect(node)) {
        DomValue::Rect {
            x,
            y,
            width,
            height,
        } => ((x + width / 2.0) as f32, (y + height / 2.0) as f32),
        other => panic!("getBoundingClientRect answered {other:?}"),
    };
    for kind in ["mousedown", "mouseup"] {
        apply(
            doc,
            DomOp::DispatchPointer {
                kind: kind.to_string(),
                client_x: x,
                client_y: y,
                button: 0,
            },
        );
    }
}

#[test]
fn a_click_comes_back_as_a_dom_event() {
    let doc = setup();
    let b = create(doc, "button", "");
    apply(doc, DomOp::SetAttribute(b, "id".into(), "go".into()));
    apply(
        doc,
        DomOp::AppendChild {
            parent: DOCUMENT,
            child: b,
        },
    );
    assert_eq!(node(apply(doc, DomOp::GetElementById("go".into()))), b);

    click_at(doc, b);

    let DomValue::Events(events) = apply(doc, DomOp::DrainEvents) else {
        panic!("DrainEvents must answer with events");
    };
    assert!(
        events.iter().any(|(n, k)| *n == b && k == "click"),
        "expected a click on the button, got {:?}",
        events
    );
    // Drained once — a second poll is empty, as an event queue must be.
    let DomValue::Events(again) = apply(doc, DomOp::DrainEvents) else {
        panic!()
    };
    assert!(
        again.is_empty(),
        "events must not be redelivered: {:?}",
        again
    );
}

fn is_open(doc: u64, dialog: u64) -> bool {
    matches!(apply(doc, DomOp::DialogOpen(dialog)), DomValue::Bool(true))
}

#[test]
fn a_dialog_is_closed_until_shown_and_closed_again_after() {
    let doc = setup();
    let dlg = create(doc, "dialog", "");
    apply(
        doc,
        DomOp::AppendChild {
            parent: DOCUMENT,
            child: dlg,
        },
    );
    assert!(!is_open(doc, dlg), "a fresh dialog must not be open");

    apply(
        doc,
        DomOp::ShowDialog {
            node: dlg,
            modal: false,
        },
    );
    assert!(is_open(doc, dlg), "show() must set the open attribute");

    apply(doc, DomOp::CloseDialog(dlg));
    assert!(!is_open(doc, dlg), "close() must clear the open attribute");
}

#[test]
fn only_a_modal_dialog_is_positioned_against_the_viewport() {
    // The UA stylesheet's own distinction: `dialog:modal` is `position:
    // fixed`, a non-modal dialog stays in flow. If both looked the same,
    // `showModal` would be `show` under another name.
    let doc = setup();
    let plain = create(doc, "dialog", "");
    let modal = create(doc, "dialog", "");
    for child in [plain, modal] {
        apply(
            doc,
            DomOp::AppendChild {
                parent: DOCUMENT,
                child,
            },
        );
    }

    apply(
        doc,
        DomOp::ShowDialog {
            node: plain,
            modal: false,
        },
    );
    apply(
        doc,
        DomOp::ShowDialog {
            node: modal,
            modal: true,
        },
    );

    assert_eq!(
        text(apply(doc, DomOp::GetStyleProperty(modal, "position".into()))),
        "fixed"
    );
    assert_ne!(
        text(apply(doc, DomOp::GetStyleProperty(plain, "position".into()))),
        "fixed",
        "a non-modal dialog stays in flow"
    );
}

/// **`removeEventListener` unsubscribes the callback it was given, and only
/// that one.**
///
/// The interesting half is identity. `Value`'s `==` compares two
/// `ObjectKind::Function`s by `chunk_index`, so every closure a factory
/// produces is "equal" to its siblings — `makeHandler(d)` in a loop, which is
/// how the calculator builds its keypad. Matching by equality would remove
/// whichever sibling happened to be first, invisibly. This pins the pointer
/// identity the spec actually means.
#[test]
fn remove_event_listener_takes_the_listener_it_was_given() {
    use std::sync::{Arc, Mutex};
    use vybe_runtime::value::{Object, ObjectKind, Value};
    use vybe_platform_web::html;

    let doc = setup();
    let button = create(doc, "button", "");
    apply(
        doc,
        DomOp::AppendChild {
            parent: DOCUMENT,
            child: button,
        },
    );
    let node = button;

    // Two DISTINCT callback objects that `==` cannot tell apart: same kind,
    // same host-function index. This is the sibling-closure shape.
    let make = || {
        let mut o = Object::new();
        o.kind = ObjectKind::HostFunction(7);
        Value::Object(Arc::new(Mutex::new(o)))
    };
    let (first, second) = (make(), make());
    // `Value::eq` is the language's own equality — the one any "find the
    // matching listener" code would reach for — and it answers TRUE here:
    // `kind_eq` compares two `HostFunction`s by index and two `Function`s by
    // `chunk_index`, so distinct closures off one factory are equal. (Rust's
    // `PartialEq` is stricter, which is why this must be asserted explicitly
    // rather than with `assert_eq!`.)
    assert!(
        first.eq(&second),
        "the two callbacks compare EQUAL — that is the trap this guards"
    );

    html::add_event_listener(doc, node, "click", first.clone());
    html::add_event_listener(doc, node, "click", second.clone());
    assert_eq!(html::listeners_for(doc, node, "click").len(), 2);

    html::remove_event_listener(doc, node, "click", &second);
    let left = html::listeners_for(doc, node, "click");
    assert_eq!(left.len(), 1, "exactly one listener should have gone");
    match (&left[0], &first) {
        (Value::Object(a), Value::Object(b)) => assert!(
            Arc::ptr_eq(a, b),
            "the WRONG listener was removed — equality matched a sibling"
        ),
        _ => panic!("listener is not an object"),
    }

    // Removing something never added removes nothing.
    html::remove_event_listener(doc, node, "click", &make());
    assert_eq!(html::listeners_for(doc, node, "click").len(), 1);

    // And the type is part of the key.
    html::remove_event_listener(doc, node, "input", &first);
    assert_eq!(html::listeners_for(doc, node, "click").len(), 1);
}

/// **`innerHTML` — set the page in one go.**
///
/// The parser, the HTML grammar and the tree-builder all existed; every entry
/// point built a NEW document, so a frontend that renders a whole page at once
/// (rather than appending controls one at a time) had nothing to call.
#[test]
fn inner_html_replaces_the_subtree_and_reads_back() {
    let doc = setup();
    let host = create(doc, "div", "");
    apply(
        doc,
        DomOp::AppendChild {
            parent: DOCUMENT,
            child: host,
        },
    );

    apply(
        doc,
        DomOp::SetInnerHtml {
            node: host,
            html: "<button>7</button><button>8</button>".into(),
        },
    );
    let kids = match apply(doc, DomOp::ChildNodes(host)) {
        DomValue::Nodes(n) => n,
        other => panic!("expected nodes, got {other:?}"),
    };
    assert_eq!(kids.len(), 2, "the fragment did not build two children");
    assert_eq!(text(apply(doc, DomOp::NodeName(kids[0]))).to_lowercase(), "button");

    // Setting again REPLACES — the spec's own wording, and the difference
    // between a page that redraws and one that grows on every render.
    apply(
        doc,
        DomOp::SetInnerHtml {
            node: host,
            html: "<span>only</span>".into(),
        },
    );
    let kids = match apply(doc, DomOp::ChildNodes(host)) {
        DomValue::Nodes(n) => n,
        other => panic!("expected nodes, got {other:?}"),
    };
    assert_eq!(kids.len(), 1, "innerHTML appended instead of replacing");

    // …and reads back as markup.
    let html = text(apply(doc, DomOp::InnerHtml(host)));
    assert!(html.contains("span"), "innerHTML read back as {html:?}");

    // Emptying is how a page clears itself.
    apply(
        doc,
        DomOp::SetInnerHtml {
            node: host,
            html: String::new(),
        },
    );
    match apply(doc, DomOp::ChildNodes(host)) {
        DomValue::Nodes(n) => assert!(n.is_empty(), "innerHTML = \"\" left {} children", n.len()),
        other => panic!("expected nodes, got {other:?}"),
    }
}

/// A fragment must NEST — the child of a child belongs inside it.
///
/// The seeded element sits at index 0 of the sink's open-element stack, so the
/// driver's depth is offset by one. Without that offset a `close_to` closed the
/// WRAPPER instead of the element inside it, and everything after the first
/// child landed as a sibling of the wrapper rather than its content.
#[test]
fn inner_html_keeps_its_nesting() {
    let doc = setup();
    let host = create(doc, "div", "");
    apply(
        doc,
        DomOp::AppendChild {
            parent: DOCUMENT,
            child: host,
        },
    );
    apply(
        doc,
        DomOp::SetInnerHtml {
            node: host,
            html: "<div id='w'><span>a</span><span>b</span><span>c</span></div>".into(),
        },
    );

    // One child of the host: the wrapper.
    let top = match apply(doc, DomOp::ChildNodes(host)) {
        DomValue::Nodes(n) => n,
        other => panic!("expected nodes, got {other:?}"),
    };
    assert_eq!(top.len(), 1, "the fragment unwrapped itself: {top:?}");

    // …and all THREE spans inside it, not just the first.
    let inner = match apply(doc, DomOp::ChildNodes(top[0])) {
        DomValue::Nodes(n) => n,
        other => panic!("expected nodes, got {other:?}"),
    };
    assert_eq!(
        inner.len(),
        3,
        "children escaped the wrapper after the first"
    );
}

/// Helper: a host `<div>` attached to the document, to build inside.
fn host(doc: u64) -> u64 {
    let host = create(doc, "div", "");
    apply(
        doc,
        DomOp::AppendChild {
            parent: DOCUMENT,
            child: host,
        },
    );
    host
}

/// The children's tags, for tests about tree SHAPE.
///
/// `localName`, not `nodeName`: `nodeName` is the HTML-uppercased qualified
/// name (DOM §4.9), so it answers `A`/`B` and would make every shape assertion
/// here about casing as well as order. `casing_is_tolerated_where_html_says…`
/// is where the uppercase rule is asserted, once.
fn tags_of(doc: u64, parent: u64) -> Vec<String> {
    match apply(doc, DomOp::ChildNodes(parent)) {
        DomValue::Nodes(children) => children
            .into_iter()
            .map(|c| text(apply(doc, DomOp::LocalName(c))))
            .collect(),
        other => panic!("expected nodes, got {other:?}"),
    }
}

#[test]
fn a_document_fragment_splices_its_children_through_the_seam() {
    install();
    let doc = new_document("t");
    let host = host(doc);

    let fragment = node(apply(doc, DomOp::CreateDocumentFragment));
    for tag in ["a", "b"] {
        let child = create(doc, tag, "");
        apply(
            doc,
            DomOp::AppendChild {
                parent: fragment,
                child,
            },
        );
    }
    apply(
        doc,
        DomOp::AppendChild {
            parent: host,
            child: fragment,
        },
    );

    // Whichever engine is installed, the fragment itself does not land.
    assert_eq!(tags_of(doc, host), vec!["a", "b"]);
}

#[test]
fn insert_adjacent_html_places_markup_without_disturbing_what_is_there() {
    install();
    let doc = new_document("t");
    let host = host(doc);
    let pivot = create(doc, "p", "");
    apply(
        doc,
        DomOp::AppendChild {
            parent: host,
            child: pivot,
        },
    );

    for (position, markup) in [
        ("beforebegin", "<a></a>"),
        ("afterend", "<s></s>"),
        ("afterbegin", "<b></b>"),
        ("beforeend", "<u></u>"),
    ] {
        apply(
            doc,
            DomOp::InsertAdjacentHtml {
                node: pivot,
                position: position.into(),
                html: markup.into(),
            },
        );
    }

    // `beforebegin`/`afterend` are the pivot's SIBLINGS, and the pivot is
    // still between them — the whole point of `insertAdjacent*` is that it
    // adds without replacing.
    assert_eq!(tags_of(doc, host), vec!["a", "p", "s"]);
    // …and the other two are its children, in order.
    assert_eq!(tags_of(doc, pivot), vec!["b", "u"]);
}

#[test]
fn outer_html_reads_the_element_and_its_setter_replaces_it() {
    install();
    let doc = new_document("t");
    let host = host(doc);
    let victim = create(doc, "p", "");
    apply(
        doc,
        DomOp::AppendChild {
            parent: host,
            child: victim,
        },
    );

    let outer = text(apply(doc, DomOp::OuterHtml(victim)));
    assert!(outer.contains("<p"), "outerHTML includes the element: {outer:?}");

    apply(
        doc,
        DomOp::SetOuterHtml {
            node: victim,
            html: "<a></a><b></b>".into(),
        },
    );
    // The element is REPLACED — not emptied, and not left beside the new
    // markup. Two nodes go in where one came out.
    assert_eq!(tags_of(doc, host), vec!["a", "b"]);
}

#[test]
fn import_node_copies_across_documents_without_inserting() {
    install();
    let source = new_document("source");
    let target = new_document("target");

    // Build `<div id="x"><span></span></div>` in the SOURCE document.
    let outer = create(source, "div", "");
    apply(source, DomOp::SetAttribute(outer, "id".into(), "x".into()));
    let inner = create(source, "span", "");
    apply(
        source,
        DomOp::AppendChild {
            parent: outer,
            child: inner,
        },
    );
    apply(
        source,
        DomOp::AppendChild {
            parent: DOCUMENT,
            child: outer,
        },
    );

    // The mechanism, asserted first: importing goes through markup, so the
    // source has to be able to describe itself.
    let described = text(apply(source, DomOp::OuterHtml(outer)));
    assert!(
        described.contains("<div") && described.contains("<span"),
        "the source node must serialise before it can be imported: {described:?}"
    );

    // Deep import: the subtree comes with it.
    let imported = apply(
        target,
        DomOp::ImportNode {
            source,
            node: outer,
            deep: true,
        },
    );
    assert!(
        matches!(imported, DomValue::Node(_)),
        "importing {described:?} produced {imported:?}"
    );
    let deep = node(imported);
    // `DIV`, not `div` — `nodeName` is the HTML-uppercased qualified name
    // (DOM §4.9). What this line is checking is that the import produced an
    // element of the right kind in the TARGET document.
    assert_eq!(text(apply(target, DomOp::NodeName(deep))), "DIV");
    assert_eq!(
        text(apply(target, DomOp::GetAttribute(deep, "id".into()))),
        "x"
    );
    assert_eq!(tags_of(target, deep), vec!["span"]);

    // …and it is DETACHED. `importNode` copies; it does not insert.
    assert!(
        matches!(apply(target, DomOp::ParentNode(deep)), DomValue::Null),
        "an imported node must have no parent until the caller inserts it"
    );

    // Shallow import: the node and nothing under it.
    let shallow = node(apply(
        target,
        DomOp::ImportNode {
            source,
            node: outer,
            deep: false,
        },
    ));
    // Positive facts FIRST: an empty child list is also what a failed import
    // looks like, so "no children" alone cannot tell the two apart.
    assert_eq!(text(apply(target, DomOp::NodeName(shallow))), "DIV");
    assert_eq!(
        text(apply(target, DomOp::GetAttribute(shallow, "id".into()))),
        "x",
        "a shallow import is still the node, attributes and all"
    );
    assert!(
        tags_of(target, shallow).is_empty(),
        "deep:false must not bring the subtree"
    );

    // The source is untouched — this is a copy, not a move.
    assert_eq!(tags_of(source, outer), vec!["span"]);
}

#[test]
fn a_fragment_accepts_children_in_any_document() {
    install();
    let _first = new_document("first");
    let second = new_document("second");

    let fragment = node(apply(second, DomOp::CreateDocumentFragment));
    let child = create(second, "div", "");
    let appended = apply(
        second,
        DomOp::AppendChild {
            parent: fragment,
            child,
        },
    );
    assert!(
        matches!(appended, DomValue::Bool(true)),
        "appending into a fragment answered {appended:?} — the parser's sink \
         treats anything but Bool(true) as a refusal and drops the node"
    );
    assert_eq!(tags_of(second, fragment), vec!["div"]);
}

#[test]
fn a_child_of_the_document_lands_in_the_body() {
    // A Document takes exactly ONE element child — `document.appendChild(<p>)`
    // is a `HierarchyRequestError` in a browser — so a caller that says "the
    // document" means the body, and every frontend here says exactly that:
    // `web:html.appendChild(DOCUMENT, control)` is how a form is built.
    //
    // webcore spelled `DOCUMENT` as `<html>` and hung the whole form beside
    // `<head>` and `<body>` instead of inside one. Nothing errored, layout ran
    // and reported its timings, and the window painted an empty page.
    let doc = setup();
    let p = create(doc, "p", "");
    apply(
        doc,
        DomOp::AppendChild {
            parent: DOCUMENT,
            child: p,
        },
    );

    let body = node(apply(doc, DomOp::QuerySelector("body".into())));
    assert_eq!(
        node(apply(doc, DomOp::ParentNode(p))),
        body,
        "a node appended to the document must be a child of the body"
    );
    assert!(
        matches!(
            apply(doc, DomOp::QuerySelector("body > p".into())),
            DomValue::Node(_)
        ),
        "the body has no element children — content went somewhere a page \
         does not render"
    );
}

#[test]
fn the_documents_own_structure_is_obeyed_where_it_is_put() {
    // The exception to the redirect: `<html>`, `<head>` and `<body>` ARE the
    // document's structure. A parser that spells one out is building the
    // skeleton, not adding content, so it is not moved into the body.
    let doc = setup();
    let extra_body = create(doc, "body", "");
    apply(
        doc,
        DomOp::AppendChild {
            parent: DOCUMENT,
            child: extra_body,
        },
    );
    assert!(
        !matches!(
            apply(doc, DomOp::QuerySelector("body > body".into())),
            DomValue::Node(_)
        ),
        "a `<body>` addressed to the document was redirected INTO the body"
    );
}

#[test]
fn the_document_is_not_an_element() {
    // DOM §4.4. `DOCUMENT` is the document node, whatever the engine below
    // uses to stand in for it — webcore has no node for it at all and answers
    // with `<html>`, which is right for reaching into the tree and wrong for
    // the two questions that ask what the node IS.
    let doc = setup();
    assert!(
        matches!(apply(doc, DomOp::NodeType(DOCUMENT)), DomValue::Number(n) if n == 9.0),
        "the document's nodeType is 9, not an element's 1"
    );
    assert_eq!(text(apply(doc, DomOp::NodeName(DOCUMENT))), "#document");
}

#[test]
fn the_documents_text_is_its_title_and_writing_it_keeps_the_tree() {
    // `textContent` replaces every child with one text node. Applied to the
    // DOCUMENT that means `<head>` and `<body>` are DELETED — which is what
    // webcore did, because it spells the document as `<html>` and applied the
    // element rule. A .NET form's caption is `Form.Text` and the form IS the
    // document, so the first line a program ran emptied its own page and every
    // control appended afterwards hung off a bodyless root.
    //
    // DOM §4.4 gives a Document null `textContent` and makes the setter a
    // no-op; `widgets` answers the title instead, and one seam cannot have
    // two answers.
    let doc = setup();
    apply(doc, DomOp::SetTextContent(DOCUMENT, "Contact Manager".into()));

    assert_eq!(
        text(apply(doc, DomOp::Title)),
        "Contact Manager",
        "the document's text is its title"
    );
    assert!(
        matches!(
            apply(doc, DomOp::QuerySelector("body".into())),
            DomValue::Node(_)
        ),
        "writing the document's text destroyed the body"
    );

    // And content appended AFTER that write still reaches the body.
    let p = create(doc, "p", "");
    apply(
        doc,
        DomOp::AppendChild {
            parent: DOCUMENT,
            child: p,
        },
    );
    assert!(
        matches!(
            apply(doc, DomOp::QuerySelector("body > p".into())),
            DomValue::Node(_)
        ),
        "the tree survived the title write but content no longer lands in it"
    );
}

#[test]
fn setting_an_elements_text_inserts_a_text_node() {
    // DOM §4.4 "string replace all": the element gets a CHILD Text node. Not a
    // decoration — an engine lays out and paints the text it finds in the tree,
    // so an implementation that keeps the string on the element instead has a
    // `textContent` and an `outerHTML` that both read back correctly while the
    // page renders a blank box. Every .NET label and button caption arrives
    // this way.
    let doc = setup();
    let label = create(doc, "label", "");
    apply(
        doc,
        DomOp::AppendChild {
            parent: DOCUMENT,
            child: label,
        },
    );
    apply(doc, DomOp::SetTextContent(label, "Name:".into()));

    let kids = match apply(doc, DomOp::ChildNodes(label)) {
        DomValue::Nodes(kids) => kids,
        other => panic!("childNodes answered {other:?}"),
    };
    assert_eq!(kids.len(), 1, "textContent must leave exactly one child");
    assert!(
        matches!(apply(doc, DomOp::NodeType(kids[0])), DomValue::Number(n) if n == 3.0),
        "the child must be a Text node (nodeType 3)"
    );
    assert_eq!(text(apply(doc, DomOp::TextContent(label))), "Name:");

    // Empty text is the spec's null case: children removed, nothing inserted.
    apply(doc, DomOp::SetTextContent(label, String::new()));
    assert!(
        matches!(apply(doc, DomOp::ChildNodes(label)), DomValue::Nodes(k) if k.is_empty()),
        "setting empty text must remove the children and insert nothing"
    );
}

#[test]
fn a_checkboxs_value_is_what_it_submits_and_checked_is_its_state() {
    // HTML §4.10.5.1.15 keeps these apart: `value` is the string a form sends
    // when the box is ticked, `checked` is whether it is ticked. Nothing about
    // one implies the other.
    //
    // The toolkit used to coerce — anything but `"false"`/`""` written to
    // `value` ticked the box — so an emitter that reached for the wrong member
    // looked correct here and rendered an EMPTY box under an engine that means
    // what HTML says. Flutter's `Checkbox(value:)` was exactly that, and this
    // is the assertion that would have caught it on either engine.
    let doc = setup();
    let cb = create(doc, "input", "checkbox");
    apply(
        doc,
        DomOp::AppendChild {
            parent: DOCUMENT,
            child: cb,
        },
    );

    // The default a form submits for a box with no `value` attribute.
    assert_eq!(text(apply(doc, DomOp::Value(cb))), "on");
    assert!(
        matches!(apply(doc, DomOp::Checked(cb)), DomValue::Bool(false)),
        "a fresh checkbox is not checked"
    );

    // Writing the submission value must NOT tick the box.
    apply(doc, DomOp::SetValue(cb, "true".into()));
    assert_eq!(text(apply(doc, DomOp::Value(cb))), "true");
    assert!(
        matches!(apply(doc, DomOp::Checked(cb)), DomValue::Bool(false)),
        "writing `value` ticked the box — that is the coercion this test exists \
         to prevent, and it hides every emitter that writes the wrong member"
    );

    // And ticking it must not disturb the submission value.
    apply(doc, DomOp::SetChecked(cb, true));
    assert!(matches!(apply(doc, DomOp::Checked(cb)), DomValue::Bool(true)));
    assert_eq!(text(apply(doc, DomOp::Value(cb))), "true");
}

#[test]
fn checkedness_is_not_the_checked_attribute() {
    // HTML §4.10.5.3. The `checked` CONTENT ATTRIBUTE is `defaultChecked` —
    // what a form reset restores to — and `input.checked` is whether the box is
    // ticked right now. Ticking a box does not rewrite the markup, and writing
    // the markup does not move a box the program has already set.
    //
    // One store for both is the same conflation `value` had, and it breaks the
    // two things the split exists for: form reset, and `getAttribute` meaning
    // "what the document says" rather than "what the user did".
    let doc = setup();
    let cb = create(doc, "input", "checkbox");
    apply(
        doc,
        DomOp::AppendChild {
            parent: DOCUMENT,
            child: cb,
        },
    );

    // The markup says nothing yet.
    assert!(
        matches!(apply(doc, DomOp::GetAttribute(cb, "checked".into())), DomValue::Null),
        "a checkbox nobody wrote markup for has no `checked` attribute"
    );

    // Ticking it is a STATE change, not a markup change.
    apply(doc, DomOp::SetChecked(cb, true));
    assert!(matches!(apply(doc, DomOp::Checked(cb)), DomValue::Bool(true)));
    assert!(
        matches!(apply(doc, DomOp::GetAttribute(cb, "checked".into())), DomValue::Null),
        "ticking the box wrote a `checked` attribute into the document — the \
         state and the markup are one store"
    );

    // And the markup is the DEFAULT: setting it must not move a box whose
    // checkedness the program has already set.
    apply(doc, DomOp::SetAttribute(cb, "checked".into(), String::new()));
    assert!(
        matches!(apply(doc, DomOp::Checked(cb)), DomValue::Bool(true)),
        "the attribute overwrote checkedness the program had already set"
    );
}

#[test]
fn casing_is_tolerated_where_html_says_and_not_where_it_does_not() {
    // **Browsers tolerate weird casing at BOTH ends** — writing and targeting —
    // for the names HTML folds, and tolerate NOTHING for the values it does
    // not. Both halves matter: fold only on write and `[DATA-Foo]` matches
    // nothing; fold an `id` and two different elements answer one query.
    let doc = setup();
    let e = create(doc, "DIV", "");
    apply(
        doc,
        DomOp::AppendChild {
            parent: DOCUMENT,
            child: e,
        },
    );
    apply(doc, DomOp::SetAttribute(e, "DATA-Foo".into(), "1".into()));
    apply(doc, DomOp::SetAttribute(e, "ID".into(), "Mixed".into()));

    // `nodeName` is the HTML-UPPERCASED qualified name (DOM §4.9); `localName`
    // is the folded one. `el.nodeName == "DIV"` is what a page writes.
    assert_eq!(text(apply(doc, DomOp::NodeName(e))), "DIV");
    assert_eq!(text(apply(doc, DomOp::LocalName(e))), "div");

    // An attribute NAME is case-insensitive, whichever spelling either side
    // used. The selector half of this matched nothing until it folded too.
    for asked in ["data-foo", "DATA-FOO", "Data-Foo"] {
        assert_eq!(
            text(apply(doc, DomOp::GetAttribute(e, asked.into()))),
            "1",
            "getAttribute({asked:?}) missed a folded attribute name"
        );
    }
    for selector in ["div", "DIV", "[data-foo]", "[DATA-foo]"] {
        assert!(
            matches!(
                apply(doc, DomOp::QuerySelector(selector.into())),
                DomValue::Node(_)
            ),
            "querySelector({selector:?}) matched nothing"
        );
    }
    assert!(
        matches!(apply(doc, DomOp::ElementsByTag("DIV".into())), DomValue::Nodes(found) if found == vec![e]),
        "getElementsByTagName folds the tag name"
    );

    // ⛔ And NOT tolerated: an `id` VALUE is case-sensitive (DOM §4.5), so a
    // page cannot reach `Mixed` by asking for `mixed`. Leniency here would let
    // two distinct elements answer one lookup.
    assert!(matches!(
        apply(doc, DomOp::GetElementById("Mixed".into())),
        DomValue::Node(_)
    ));
    assert!(
        matches!(apply(doc, DomOp::GetElementById("mixed".into())), DomValue::Null),
        "an id lookup folded the case — ids are case-SENSITIVE"
    );
}

// ─── Interaction: a click has to MOVE the control's value ───────────────────
//
// Everything above drives the DOM programmatically. These drive the POINTER,
// which is the half a user exercises and the half that was reported broken:
// a list box that would not select, a dropdown that would not open, a
// scrollbar that would not move.

/// Click a point in document coordinates, both edges, as the shell does.
///
/// The companion of `click_at`, which can only reach a node's CENTRE. A list
/// box's rows are all inside one element, so selecting the third one is a
/// question about WHERE in the box the pointer went — a centre click cannot
/// ask it.
fn click_point(doc: u64, node: u64, x: f32, y: f32) {
    // Same reason as `click_at`: hit-testing needs layout, and asking for a
    // resolved value is the page-level way to force it.
    apply(doc, DomOp::ComputedStyleProperty(node, "width".into()));
    for kind in ["mousedown", "mouseup"] {
        apply(
            doc,
            DomOp::DispatchPointer {
                kind: kind.to_string(),
                client_x: x,
                client_y: y,
                button: 0,
            },
        );
    }
}

fn rect(doc: u64, node: u64) -> (f32, f32, f32, f32) {
    match apply(doc, DomOp::BoundingClientRect(node)) {
        DomValue::Rect { x, y, width, height } => {
            (x as f32, y as f32, width as f32, height as f32)
        }
        other => panic!("getBoundingClientRect answered {other:?}"),
    }
}

fn selected(doc: u64, node: u64) -> i32 {
    match apply(doc, DomOp::SelectedIndex(node)) {
        DomValue::Number(n) => n as i32,
        other => panic!("selectedIndex answered {other:?}"),
    }
}

/// A LIST BOX — `<select>` with `size` above one (HTML §4.10.7) — shows its
/// options as rows, and clicking a row selects THAT row.
///
/// This is the control a WinForms `ListBox` and a VCL `TListBox` become, so
/// "the list box will not select" is exactly this assertion.
///
/// ⛔ The ONE test in this file that is not engine-agnostic. Aiming at a row
/// needs the row HEIGHT, and the seam has no operation that answers it — a
/// row is not a node, so it has no rect to ask for. Taking the metric from
/// webcore is what keeps the aim honest rather than quartering the box, and
/// that is a webcore symbol, so the test compiles only with that engine. The
/// rest of the file states the WHATWG contract and runs against either.
#[cfg(feature = "engine-webcore")]
#[test]
fn clicking_a_list_box_row_selects_that_row() {
    let doc = setup();
    let s = create(doc, "select", "");
    apply(doc, DomOp::SetAttribute(s, "size".into(), "4".into()));
    apply(doc, DomOp::AppendChild { parent: DOCUMENT, child: s });
    for label in ["one", "two", "three", "four"] {
        apply(doc, DomOp::AddItem(s, label.into()));
    }
    // ⛔ −1, not 0. The selectedness setting algorithm auto-selects the first
    // option only at a display size of 1, so a LIST BOX rests with nothing
    // selected. Asserting 0 here applied the drop-down's rule to a list box.
    assert_eq!(selected(doc, s), -1, "a fresh list box has no selection");

    let (x, y, w, h) = rect(doc, s);
    assert!(w > 0.0 && h > 0.0, "the list box has no geometry to click: {w}x{h}");
    // The THIRD of four rows, off the ENGINE's own row metric. Quartering the
    // box happens to agree only while the box is exactly four rows tall — it
    // is not what decides which row a click lands on.
    let row_h = webcore::html::forms::list_box_row_height(16.0);
    click_point(doc, s, x + w / 2.0, y + webcore::html::forms::LIST_BOX_PADDING + row_h * 2.5);

    assert_eq!(
        selected(doc, s),
        2,
        "clicking the third row must select the third option"
    );
}

/// A DROPDOWN is the same element without `size`, and clicking it must open
/// its list rather than do nothing.
///
/// Openness is not in the DOM — HTML gives `<select>` no `open` IDL member,
/// unlike `<dialog>` and `<details>` — so what is asserted is the observable
/// consequence: the click reaches the control and comes back as a DOM event on
/// it. A dropdown that "does nothing" fails here.
///
/// ⛔ Ignored under the WIDGETS engine, and only there: it fires no `click`
/// when a `<select>` is activated, so a program with a handler on its combo
/// box hears nothing. webcore answers this correctly and runs the test for
/// real. The assertion is the WHATWG contract for both — this states which
/// engine does not meet it yet rather than deleting the question.
#[cfg_attr(
    not(feature = "engine-webcore"),
    ignore = "widgets fires no click when a <select> is activated"
)]
#[test]
fn clicking_a_dropdown_reaches_the_control() {
    let doc = setup();
    let s = create(doc, "select", "");
    apply(doc, DomOp::AppendChild { parent: DOCUMENT, child: s });
    for label in ["alpha", "beta"] {
        apply(doc, DomOp::AddItem(s, label.into()));
    }
    // Drain what building the tree may have queued, so what is asserted below
    // is this click and not a leftover.
    apply(doc, DomOp::DrainEvents);

    click_at(doc, s);

    let DomValue::Events(events) = apply(doc, DomOp::DrainEvents) else {
        panic!("DrainEvents must answer with events");
    };
    assert!(
        events.iter().any(|(n, k)| *n == s && k == "click"),
        "a click on a dropdown never reached it, got {events:?}"
    );
    // ⛔ EXACTLY one. `any()` cannot tell one click from two, and two is what a
    // control gets when a generic click path and a per-control one both fire:
    // a handler that counts, toggles or appends would do it twice for one press.
    assert_eq!(
        events.iter().filter(|(_, k)| k == "click").count(),
        1,
        "one press is one click, got {events:?}"
    );
}

/// One press is one `click`, whatever kind of element it lands on.
///
/// The generic path and the per-control paths both live in the mouse-up arm,
/// and nothing structural stops both from firing. This counts.
#[cfg_attr(
    not(feature = "engine-webcore"),
    ignore = "widgets fires click only for form controls — a <div> listener never runs"
)]
#[test]
fn one_press_is_one_click_on_every_kind_of_element() {
    let doc = setup();
    for (tag, input_type) in [("div", ""), ("button", ""), ("input", "checkbox"), ("select", "")] {
        let e = create(doc, tag, input_type);
        apply(doc, DomOp::AppendChild { parent: DOCUMENT, child: e });
        apply(doc, DomOp::SetTextContent(e, "xx".into()));
        apply(doc, DomOp::DrainEvents);

        click_at(doc, e);

        let DomValue::Events(events) = apply(doc, DomOp::DrainEvents) else {
            panic!("DrainEvents must answer with events");
        };
        let clicks = events.iter().filter(|(_, k)| k == "click").count();
        assert_eq!(clicks, 1, "<{tag}> reported {clicks} clicks for one press: {events:?}");
        apply(doc, DomOp::RemoveChild { parent: DOCUMENT, child: e });
    }
}

/// A SCROLLBAR is `<input type=range>` — a value in a range — and clicking
/// along its track moves the value there.
///
/// This is what `HScrollBar`/`VScrollBar` and `TrackBar` become, so "the
/// scrollbars do not work" is this assertion for the horizontal case.
#[test]
fn clicking_a_range_track_moves_its_value() {
    let doc = setup();
    let r = create(doc, "input", "range");
    apply(doc, DomOp::SetAttribute(r, "min".into(), "0".into()));
    apply(doc, DomOp::SetAttribute(r, "max".into(), "100".into()));
    apply(doc, DomOp::SetValue(r, "0".into()));
    apply(doc, DomOp::AppendChild { parent: DOCUMENT, child: r });

    let (x, y, w, h) = rect(doc, r);
    assert!(w > 0.0 && h > 0.0, "the range has no geometry to click: {w}x{h}");
    // Three quarters along the track. Not the exact end, because a thumb has
    // width and the last pixel is not reachable by its centre.
    click_point(doc, r, x + w * 0.75, y + h / 2.0);

    let after: f32 = text(apply(doc, DomOp::Value(r))).parse().unwrap_or(-1.0);
    assert!(
        after > 50.0,
        "clicking three quarters along the track left the value at {after}"
    );
}

/// **Every element the seam inserts has to OCCUPY SPACE.**
///
/// A box with no geometry cannot be hit-tested, so a control that lays out to
/// 0x0 is unclickable however well its event path works — which makes this the
/// question to ask before any "the click did nothing" report is believed.
#[test]
fn an_inserted_control_has_geometry() {
    let doc = setup();
    for (tag, input_type) in [
        ("button", ""),
        ("input", "text"),
        ("input", "range"),
        ("select", ""),
        ("div", ""),
    ] {
        let e = create(doc, tag, input_type);
        apply(doc, DomOp::AppendChild { parent: DOCUMENT, child: e });
        apply(doc, DomOp::SetTextContent(e, "xx".into()));
        let (_, _, w, h) = rect(doc, e);
        assert!(
            w > 0.0 && h > 0.0,
            "<{tag} type={input_type}> laid out to {w}x{h} — nothing can click it"
        );
    }
}

/// `el.innerHTML = ""` EMPTIES the element — including children that were
/// appended one at a time rather than parsed in.
///
/// This is how a page clears itself before redrawing, and a control that
/// rebuilds its own contents does exactly that. If the clear silently keeps
/// what was appended, every redraw stacks another copy on top of the last —
/// which looks like a rendering bug and is a DOM one.
#[cfg_attr(
    not(feature = "engine-webcore"),
    ignore = "widgets does not update its selector index when a subtree is removed"
)]
#[test]
fn clearing_inner_html_empties_an_element_that_was_built_by_appending() {
    let doc = setup();
    let box_ = create(doc, "div", "");
    apply(doc, DomOp::AppendChild { parent: DOCUMENT, child: box_ });
    for _ in 0..3 {
        let kid = create(doc, "span", "");
        apply(doc, DomOp::AppendChild { parent: box_, child: kid });
        // ⛔ NESTED, not three flat spans. A control that rebuilds itself has a
        // subtree, not a row of leaves, and "remove the children" has to mean
        // the whole of each one.
        let grandkid = create(doc, "b", "");
        apply(doc, DomOp::SetAttribute(grandkid, "class".into(), "gone".into()));
        apply(doc, DomOp::SetTextContent(grandkid, "x".into()));
        apply(doc, DomOp::AppendChild { parent: kid, child: grandkid });
    }
    let before = match apply(doc, DomOp::ChildNodes(box_)) {
        DomValue::Nodes(n) => n.len(),
        other => panic!("expected children, got {other:?}"),
    };
    assert_eq!(before, 3, "the fixture must have children to clear");

    apply(doc, DomOp::SetInnerHtml { node: box_, html: String::new() });

    let after = match apply(doc, DomOp::ChildNodes(box_)) {
        DomValue::Nodes(n) => n.len(),
        other => panic!("expected children, got {other:?}"),
    };
    assert_eq!(after, 0, "innerHTML = \"\" must leave the element empty");

    // ⛔ AND THE SELECTOR ENGINE MUST AGREE. `childNodes` reading empty while
    // `querySelectorAll` still matches the removed subtree is two stores
    // disagreeing about one document, and it is invisible until something
    // rebuilds itself and appears to double.
    assert!(
        nodes_matching(doc, ".gone").is_empty(),
        "a removed subtree must not still match a selector"
    );
}

fn nodes_matching(doc: u64, selector: &str) -> Vec<u64> {
    match apply(doc, DomOp::QuerySelectorAll(selector.to_string())) {
        DomValue::Nodes(n) => n,
        other => panic!("expected a node list, got {other:?}"),
    }
}

/// A control BUILT BEFORE IT IS APPENDED, then cleared and rebuilt — the exact
/// lifecycle of any control that fills itself in at construction.
///
/// `New MonthCalendar()` builds its whole month while the element is still
/// detached, and only then does `Controls.Add` put it in the document. Every
/// later redraw clears it and builds again. If the clear misses the copy the
/// selector engine walks, each redraw appears to ADD a calendar rather than
/// replace one — which reads as a rendering bug and is a store-consistency one.
#[cfg_attr(
    not(feature = "engine-webcore"),
    ignore = "widgets does not update its selector index when a subtree is removed"
)]
#[test]
fn a_subtree_built_while_detached_can_still_be_cleared_once_appended() {
    let doc = setup();
    let host = create(doc, "div", "");

    // Built DETACHED, exactly as a control fills itself in at construction.
    for _ in 0..2 {
        let part = create(doc, "section", "");
        apply(doc, DomOp::SetAttribute(part, "class".into(), "first-pass".into()));
        let leaf = create(doc, "b", "");
        apply(doc, DomOp::SetAttribute(leaf, "class".into(), "first-leaf".into()));
        apply(doc, DomOp::AppendChild { parent: part, child: leaf });
        apply(doc, DomOp::AppendChild { parent: host, child: part });
    }
    // …and only now put in the document.
    apply(doc, DomOp::AppendChild { parent: DOCUMENT, child: host });
    assert_eq!(nodes_matching(doc, ".first-leaf").len(), 2, "the fixture must be built");

    // The redraw: clear, then build again.
    apply(doc, DomOp::SetTextContent(host, String::new()));
    for _ in 0..2 {
        let part = create(doc, "section", "");
        apply(doc, DomOp::SetAttribute(part, "class".into(), "second-pass".into()));
        apply(doc, DomOp::AppendChild { parent: host, child: part });
    }

    assert!(
        nodes_matching(doc, ".first-pass").is_empty(),
        "the cleared subtree must not still match a selector"
    );
    assert!(
        nodes_matching(doc, ".first-leaf").is_empty(),
        "nor must anything inside it"
    );
    assert_eq!(nodes_matching(doc, ".second-pass").len(), 2, "the rebuild is what is there now");
}
