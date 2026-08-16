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

fn setup() -> u64 {
    vybe_platform_web::engine_widgets::install();
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
    apply(doc, DomOp::SetValue(s, "1".into()));
    assert_eq!(text(apply(doc, DomOp::Value(s))), "1");
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
    assert_eq!(
        text(apply(doc, DomOp::GetStyleProperty(b, "top".into()))),
        "16px"
    );
}

#[test]
fn two_documents_are_two_trees() {
    vybe_platform_web::engine_widgets::install();
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

#[test]
fn a_click_comes_back_as_a_dom_event() {
    use vybe_platform_web::engine_widgets::with_document;
    use vybe_widgets::layout::{MouseButton, MouseEvent, MouseEventKind, PanelWidget};

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
