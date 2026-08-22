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
///     cargo test -p vybe_platform_web --features gui             # vybe_widgets
///     cargo test -p vybe_platform_web --features engine-htmlbox  # htmlbox
fn install() {
    #[cfg(feature = "engine-htmlbox")]
    vybe_platform_web::engine_htmlbox::install();
    #[cfg(not(feature = "engine-htmlbox"))]
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
    // `getPropertyValue` silently rewrote authored units. htmlbox always
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
/// The one place an engine difference is unavoidable: injecting OS-level input
/// is not a WHATWG operation — a browser receives it from the platform and no
/// page script does this. Everything the test then ASSERTS is standard.
#[cfg(not(feature = "engine-htmlbox"))]
fn click_at(doc: u64, _node: u64) {
    use vybe_platform_web::engine_widgets::with_document;
    use vybe_widgets::layout::{MouseButton, MouseEvent, MouseEventKind, PanelWidget};

    let press = MouseEvent {
        kind: MouseEventKind::Press(MouseButton::Left),
        x: 10.0,
        y: 10.0,
        cmd: false,
        shift: false,
        alt: false,
    };
    with_document(doc, |d| {
        let form = d.form_mut();
        form.handle_mouse(&press);
        form.handle_mouse(&MouseEvent {
            kind: MouseEventKind::Release(MouseButton::Left),
            ..press
        });
    });
}

#[cfg(feature = "engine-htmlbox")]
fn click_at(doc: u64, node: u64) {
    use rhtmledit::dom::HtmlEventType;
    use rhtmledit::layout::LayoutEngine;
    use vybe_platform_web::engine_htmlbox::with_document;

    with_document(doc, |d| {
        // Hit-testing needs boxes to have rects, and a document that has never
        // been laid out has none.
        LayoutEngine::new().layout(d, 1024.0);
        let pt = d
            .get_bounding_client_rect(node as u32)
            .map(|r| (r.x + r.w / 2.0, r.y + r.h / 2.0))
            .unwrap_or((0.0, 0.0));
        d.process_mouse_event(HtmlEventType::MouseDown, pt, 0);
        d.process_mouse_event(HtmlEventType::MouseUp, pt, 0);
        d.process_mouse_event(HtmlEventType::Click, pt, 0);
    });
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

fn tags_of(doc: u64, parent: u64) -> Vec<String> {
    match apply(doc, DomOp::ChildNodes(parent)) {
        DomValue::Nodes(children) => children
            .into_iter()
            .map(|c| text(apply(doc, DomOp::NodeName(c))))
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
    assert_eq!(text(apply(target, DomOp::NodeName(deep))), "div");
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
