//! Plugin init must leave the selected browser engine and canvas backend paired.
//!
//! This is the path `vybex` uses. A lower-level seam test can install matching
//! pieces manually and still miss a plugin init that overwrites only one slot.

#![cfg(feature = "engine-webcore")]

use vybe_platform_web::canvas_backend::{Op2D, apply as paint, backend};
use vybe_platform_web::engine::{DOCUMENT, DomOp, DomValue, apply};
use vybe_platform_web::engine_select::{self, Engine};
use vybe_runtime::{Framework, Plugin as _};

fn node(v: DomValue) -> u64 {
    match v {
        DomValue::Node(n) => n,
        other => panic!("expected node, got {other:?}"),
    }
}

#[test]
fn plugin_init_keeps_webcore_canvas_backend_with_webcore_engine() {
    engine_select::choose(Engine::WebCore);

    let mut vm = vybe_runtime::VM::new();
    let mut fw = Framework::with_vm(&mut vm);
    vybe_platform_web::Plugin.init(&mut fw);

    assert_eq!(engine_select::live(), Some(Engine::WebCore));

    let doc = vybe_platform_web::html::active_document();
    let canvas = node(apply(
        doc,
        DomOp::CreateElement {
            tag: "canvas".into(),
            input_type: String::new(),
        },
    ));
    apply(doc, DomOp::SetAttribute(canvas, "width".into(), "80".into()));
    apply(doc, DomOp::SetAttribute(canvas, "height".into(), "40".into()));
    apply(
        doc,
        DomOp::AppendChild {
            parent: DOCUMENT,
            child: canvas,
        },
    );

    let target = format!("d{doc}:n{canvas}");
    backend().expect("canvas backend installed").ensure(&target);
    paint(&target, Op2D::SetFillStyle(255, 0, 0, 255));
    paint(&target, Op2D::FillRect(0.0, 0.0, 80.0, 40.0));

    let mut pixmap = widgets::Pixmap::new(160, 80).expect("pixmap");
    assert!(
        vybe_platform_web::present::render(doc, &mut pixmap, 1.0),
        "webcore refused to paint a document with a canvas"
    );
    assert!(
        pixmap
            .data()
            .chunks_exact(4)
            .any(|px| px[0] > 200 && px[1] < 50 && px[2] < 50 && px[3] > 200),
        "canvas draw did not reach the webcore-rendered frame"
    );
}
