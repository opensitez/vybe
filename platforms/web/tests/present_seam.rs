//! "Is there a UI in this document, and can it be painted?" — asked of
//! whichever engine is live.
//!
//! Its own binary because the answer depends on the PROCESS-WIDE engine
//! choice: `present` reads `engine_select::live()`, which a test sharing a
//! binary with another engine's `install()` would be racing.
//!
//!     cargo test -p vybe_platform_web --features gui            --test present_seam
//!     cargo test -p vybe_platform_web --features engine-htmlbox --test present_seam

use vybe_platform_web::engine::{DOCUMENT, DomOp, DomValue, apply, new_document};
use vybe_platform_web::engine_select::{self, Engine};

fn setup() -> u64 {
    // Through `engine_select`, not the engine's own `install()`: `present`
    // asks which engine is LIVE, and only this sets that.
    #[cfg(feature = "engine-htmlbox")]
    engine_select::choose(Engine::HtmlBox);
    #[cfg(not(feature = "engine-htmlbox"))]
    engine_select::choose(Engine::Widgets);
    engine_select::install();
    new_document("test")
}

#[test]
fn an_empty_document_has_no_content_and_a_populated_one_does() {
    let doc = setup();
    assert!(
        !vybe_platform_web::present::has_content(doc),
        "a document nothing was built into reported content — `active_document` \
         CREATES one on first touch, so its existence must not count as a UI"
    );

    // Exactly what a .NET form does: `document.body` is the form, and every
    // control is appended to it. `body` answers `DOCUMENT`, so this is the
    // path that has to work.
    let button = match apply(
        doc,
        DomOp::CreateElement {
            tag: "button".into(),
            input_type: String::new(),
        },
    ) {
        DomValue::Node(n) => n,
        other => panic!("createElement answered {other:?}"),
    };
    apply(
        doc,
        DomOp::AppendChild {
            parent: DOCUMENT,
            child: button,
        },
    );

    assert!(
        vybe_platform_web::present::has_content(doc),
        "a document with a control in it reported nothing to paint — which is \
         what `--capture` reports as `no live document to capture`"
    );
}

#[test]
fn a_frame_is_painted_from_the_live_engine() {
    let doc = setup();
    let button = match apply(
        doc,
        DomOp::CreateElement {
            tag: "button".into(),
            input_type: String::new(),
        },
    ) {
        DomValue::Node(n) => n,
        other => panic!("createElement answered {other:?}"),
    };
    apply(
        doc,
        DomOp::AppendChild {
            parent: DOCUMENT,
            child: button,
        },
    );
    apply(
        doc,
        DomOp::SetStyleProperty(button, "background-color".into(), "#ff0000".into()),
    );

    let mut pixmap = vybe_widgets::Pixmap::new(200, 100).expect("pixmap");
    assert!(
        vybe_platform_web::present::render(doc, &mut pixmap, 1.0),
        "the live engine refused to paint a document that has content"
    );
    // A frame that is entirely transparent is a frame nothing drew into. The
    // engines disagree about colours and geometry, deliberately; they cannot
    // disagree about whether they painted at all.
    assert!(
        pixmap.data().chunks_exact(4).any(|px| px[3] != 0),
        "the frame came back blank — the renderer painted a tree that is not \
         the one the controls went into"
    );
}
