//! The `web:canvas` seam end to end: a page finds a `<canvas>` by id, asks it
//! for a context, and draws.
//!
//! Drives the backend trait directly rather than through the VM, so what is
//! under test is the surface's contract with the engine — the part a real
//! browser backend would have to satisfy too.
//!
//! Every assertion below is the WHATWG contract, so the same file runs against
//! either engine and there is no second copy to drift:
//!
//!     cargo test -p vybe_platform_web --features gui             # widgets
//!     cargo test -p vybe_platform_web --features engine-webcore  # webcore

use vybe_platform_web::canvas_backend::{
    Op2D, Query2D, Query2DValue, StringAttribute, apply as paint, backend, query,
};
use vybe_platform_web::engine::{DOCUMENT, DomOp, DomValue, apply};

/// `measureText(text).width` on `target`, which must exist.
fn measured(target: &str, text: &str) -> f32 {
    match query(target, Query2D::MeasureText(text.to_string())) {
        Query2DValue::Metrics(m) => m.width,
        other => panic!("expected metrics, got {other:?}"),
    }
}

fn node(v: DomValue) -> u64 {
    match v {
        DomValue::Node(n) => n,
        other => panic!("expected a node, got {other:?}"),
    }
}

/// Install whichever browser this build selected — engine AND painter, which
/// have to be the same one. They are two `install()` calls because they are two
/// traits, and a build that swapped only the first would deliver every paint op
/// to a document that does not contain the node it names.
fn install() {
    #[cfg(feature = "engine-webcore")]
    {
        vybe_platform_web::engine_webcore::install();
        vybe_platform_web::canvas_backend_webcore::install();
    }
    #[cfg(not(feature = "engine-webcore"))]
    {
        vybe_platform_web::engine_widgets::install();
        vybe_platform_web::canvas_backend_widgets::install();
    }
}

/// Add a `<canvas id="{id}">` to the page, and answer the target string an
/// element-bound context carries — which is what `getContext` derives from the
/// node it was handed.
///
/// The AMBIENT document, not a fresh one: a painter resolves the surface
/// through the document a page is in, so a canvas made anywhere else is a node
/// the painter cannot see. That is not a limitation to work around in a test —
/// it is the thing being tested, and reaching for `new_document` here would
/// have hidden it.
///
/// Each test uses its own `id` because they share that one page.
fn canvas_on_the_page(id: &str) -> (u64, String) {
    install();
    let doc = vybe_platform_web::html::active_document();
    let canvas = node(apply(
        doc,
        DomOp::CreateElement {
            tag: "canvas".into(),
            input_type: String::new(),
        },
    ));
    apply(doc, DomOp::SetAttribute(canvas, "id".into(), id.into()));
    apply(
        doc,
        DomOp::AppendChild {
            parent: DOCUMENT,
            child: canvas,
        },
    );
    (doc, format!("n{canvas}"))
}

#[test]
fn a_canvas_found_by_id_is_the_one_that_gets_the_context() {
    // `getElementById` then `getContext` — the two hops a page makes, and the
    // reason the seam takes a node rather than a name.
    let (doc, target) = canvas_on_the_page("found-by-id");
    let found = node(apply(doc, DomOp::GetElementById("found-by-id".into())));
    assert_eq!(
        target,
        format!("n{found}"),
        "the element the page finds must be the one the context binds to"
    );
    backend().expect("a painter is installed").ensure(&target);
}

#[test]
fn the_drawing_state_survives_between_ops() {
    // Each op crosses the seam on its own, so a backend that rebuilt the
    // context per call would lose the font before the measurement that depends
    // on it. This is the smallest thing that catches it, because `measureText`
    // is the only answer the seam can carry back.
    let (_doc, target) = canvas_on_the_page("state-survives");
    let b = backend().expect("a painter is installed");
    b.ensure(&target);

    paint(
        &target,
        Op2D::SetFont {
            family: "sans-serif".into(),
            size: 12.0,
            bold: false,
            italic: false,
        },
    );
    let small = measured(&target, "HHHHHHHH");

    paint(
        &target,
        Op2D::SetFont {
            family: "sans-serif".into(),
            size: 48.0,
            bold: false,
            italic: false,
        },
    );
    let large = measured(&target, "HHHHHHHH");

    assert!(small > 0.0, "measured nothing at 12px");
    assert!(
        large > small * 2.0,
        "the font set by an earlier op did not survive: 12px measured {small}, 48px measured {large}"
    );
}

#[test]
fn drawing_at_a_target_that_names_nothing_is_not_an_error() {
    // A paint op with nowhere to land is invisible by design — a page with no
    // renderer attached draws nothing, it does not fail. A QUESTION is the
    // exception, and says so by answering `Absent`.
    install();
    paint("n999999", Op2D::FillRect(0.0, 0.0, 10.0, 10.0));
    assert!(
        matches!(
            query("n999999", Query2D::MeasureText("x".into())),
            Query2DValue::Absent
        ),
        "a measurement with no surface must be absent, not a plausible number"
    );
}

#[test]
fn a_full_drawing_sequence_crosses_the_seam() {
    // The ops a page actually emits, in the order it emits them: set state,
    // build a path, fill it, draw text. Nothing here can read pixels back —
    // the seam has no return path for `getImageData` — so what this pins is
    // that every arm is reachable and none of them panics on the way through.
    let (_doc, target) = canvas_on_the_page("full-sequence");
    let b = backend().expect("a painter is installed");
    b.ensure(&target);

    for op in [
        Op2D::Save,
        Op2D::SetFillStyle(255, 0, 0, 255),
        Op2D::SetStrokeStyle(0, 0, 255, 255),
        Op2D::SetLineWidth(3.0),
        Op2D::SetLineDash(vec![4.0, 2.0]),
        Op2D::SetLineCap("round".into()),
        Op2D::SetLineJoin("bevel".into()),
        Op2D::SetMiterLimit(8.0),
        Op2D::SetLineDashOffset(1.0),
        Op2D::SetGlobalAlpha(0.5),
        Op2D::SetImageSmoothing(true),
        Op2D::Translate(5.0, 5.0),
        Op2D::Scale(2.0, 2.0),
        Op2D::Rotate(0.25),
        Op2D::Transform(1.0, 0.0, 0.0, 1.0, 1.0, 1.0),
        Op2D::ResetTransform,
        Op2D::BeginPath,
        Op2D::MoveTo(0.0, 0.0),
        Op2D::LineTo(20.0, 0.0),
        Op2D::QuadraticCurveTo(25.0, 5.0, 20.0, 10.0),
        Op2D::BezierCurveTo(15.0, 15.0, 5.0, 15.0, 0.0, 10.0),
        Op2D::Arc(10.0, 10.0, 4.0, 0.0, 6.28, false),
        Op2D::Ellipse(10.0, 10.0, 6.0, 3.0),
        Op2D::Rect(1.0, 1.0, 8.0, 8.0),
        Op2D::ClosePath,
        Op2D::Fill,
        Op2D::Stroke,
        Op2D::Clip,
        Op2D::FillRect(0.0, 0.0, 10.0, 10.0),
        Op2D::StrokeRect(2.0, 2.0, 6.0, 6.0),
        Op2D::ClearRect(3.0, 3.0, 2.0, 2.0),
        Op2D::SetTextAlign("center".into()),
        Op2D::SetTextBaseline("middle".into()),
        Op2D::FillText("hi".into(), 4.0, 4.0),
        Op2D::StrokeText("hi".into(), 4.0, 12.0),
        Op2D::PutImageData {
            pixels: vec![255u8; 4 * 4 * 4],
            width: 4,
            height: 4,
            dx: 0.0,
            dy: 0.0,
        },
        Op2D::DrawImageRgba {
            pixels: vec![128u8; 2 * 2 * 4],
            width: 2,
            height: 2,
            dx: 0.0,
            dy: 0.0,
            dw: 4.0,
            dh: 4.0,
        },
        Op2D::DrawImagePaletted {
            indices: vec![0, 1, 1, 0],
            palette: vec![255, 0, 0, 0, 255, 0],
            width: 2,
            height: 2,
            dx: 0.0,
            dy: 0.0,
            dw: 4.0,
            dh: 4.0,
        },
        Op2D::Restore,
    ] {
        paint(&target, op);
    }

    // Still alive and still answering afterwards.
    assert!(measured(&target, "x") >= 0.0);
}

// ── The half that ASKS ──────────────────────────────────────────────────────
//
// These are the members that had no wire format at all until the seam grew a
// query channel. Every one runs against BOTH engines from this one file, which
// is what "exactly the same API" is actually enforced by — the trait being
// identical is a fact about two crates; this is a fact about what a page gets.


fn text(v: Query2DValue) -> String {
    match v {
        Query2DValue::Text(t) => t,
        other => panic!("expected text, got {other:?}"),
    }
}

#[test]
fn a_measurement_carries_all_twelve_metrics() {
    // The seam used to carry `width` alone: the engine computed the other
    // eleven and they were dropped on the way out.
    let (_doc, target) = canvas_on_the_page("metrics");
    paint(
        &target,
        Op2D::SetFontCss("48px sans-serif".into()),
    );
    let Query2DValue::Metrics(m) = query(&target, Query2D::MeasureText("Hg".into())) else {
        panic!("a canvas that exists can be measured");
    };
    assert!(m.width > 0.0, "no width");
    assert!(
        m.font_bounding_box_ascent > 0.0,
        "the font box has no ascent — the metrics are not being filled in"
    );
    assert!(
        m.actual_bounding_box_ascent > 0.0,
        "the glyph box has no ascent"
    );
    // `Hg` has a descender, so the glyph box must reach below the baseline.
    assert!(
        m.actual_bounding_box_descent > 0.0,
        "a string with a descender measured none"
    );
}

#[test]
fn a_page_can_read_back_what_it_set() {
    // §4.12.5 requires every string attribute to serialize. A canvas that
    // accepts a value and cannot report it is half an attribute.
    let (_doc, target) = canvas_on_the_page("read-back");
    paint(&target, Op2D::SetFillStyleCss("#3366cc".into()));
    paint(&target, Op2D::SetFontCss("32px serif".into()));
    paint(&target, Op2D::SetGlobalCompositeOperation("multiply".into()));
    paint(&target, Op2D::SetDirection("rtl".into()));

    let get = |a| text(query(&target, Query2D::GetStringAttribute(a)));
    assert_eq!(get(StringAttribute::FillStyle), "#3366cc");
    assert_eq!(get(StringAttribute::Font), "32px serif");
    assert_eq!(get(StringAttribute::GlobalCompositeOperation), "multiply");
    assert_eq!(get(StringAttribute::Direction), "rtl");
}

#[test]
fn a_colour_serializes_the_way_the_spec_says_and_not_as_written() {
    // `fillStyle = "red"` reads back `"#ff0000"`: the attribute returns the
    // SERIALIZATION of the colour, not the text the page typed. Echoing the
    // input would look more faithful and be wrong.
    let (_doc, target) = canvas_on_the_page("serialize");
    paint(&target, Op2D::SetFillStyleCss("red".into()));
    let get = || text(query(&target, Query2D::GetStringAttribute(StringAttribute::FillStyle)));
    assert_eq!(get(), "#ff0000");

    // Translucent goes to `rgba(...)`, because hex cannot carry alpha.
    paint(&target, Op2D::SetFillStyleCss("rgba(0, 128, 255, 0.5)".into()));
    let out = get();
    assert!(
        out.starts_with("rgba(0, 128, 255,"),
        "expected an rgba serialization, got {out}"
    );
}

#[test]
fn an_unparseable_value_leaves_the_attribute_alone() {
    // §4.12.5: a value that cannot be parsed is IGNORED. Resetting to a default
    // would be a silent change to something the page never asked to change.
    let (_doc, target) = canvas_on_the_page("bad-value");
    paint(&target, Op2D::SetFillStyleCss("#123456".into()));
    paint(&target, Op2D::SetFillStyleCss("not-a-colour".into()));
    assert_eq!(
        text(query(
            &target,
            Query2D::GetStringAttribute(StringAttribute::FillStyle)
        )),
        "#123456",
        "a bad assignment overwrote a good value"
    );
}

#[test]
fn what_was_drawn_can_be_read_back_through_the_seam() {
    // The query that could not exist before: `apply` returns nothing, so there
    // was no shape for an answer to travel in.
    let (_doc, target) = canvas_on_the_page("pixels");
    paint(&target, Op2D::SetFillStyleCss("#ff8000".into()));
    paint(&target, Op2D::FillRect(0.0, 0.0, 8.0, 8.0));
    let Query2DValue::Pixels { data, width, height } = query(
        &target,
        Query2D::GetImageData {
            sx: 0,
            sy: 0,
            sw: 4,
            sh: 4,
        },
    ) else {
        panic!("a canvas has pixels to hand back");
    };
    assert_eq!((width, height), (4, 4));
    // STRAIGHT RGBA, not premultiplied — §4.12.5.
    assert_eq!(&data[..4], &[255, 128, 0, 255]);
}

#[test]
fn a_hit_test_crosses_the_seam() {
    let (_doc, target) = canvas_on_the_page("hit-test");
    paint(&target, Op2D::BeginPath);
    paint(&target, Op2D::Rect(10.0, 10.0, 50.0, 50.0));
    let inside = query(
        &target,
        Query2D::IsPointInPath {
            x: 30.0,
            y: 30.0,
            rule: "nonzero".into(),
        },
    );
    let outside = query(
        &target,
        Query2D::IsPointInPath {
            x: 5.0,
            y: 5.0,
            rule: "nonzero".into(),
        },
    );
    assert!(matches!(inside, Query2DValue::Bool(true)));
    assert!(matches!(outside, Query2DValue::Bool(false)));
}

#[test]
fn the_dash_list_reads_back_normalized() {
    // §4.12.5: an ODD-length list is concatenated with itself, because a dash
    // pattern is read in on/off pairs. `setLineDash([5])` reads back `[5, 5]`.
    let (_doc, target) = canvas_on_the_page("dash");
    paint(&target, Op2D::SetLineDash(vec![5.0]));
    let Query2DValue::Floats(dash) = query(&target, Query2D::GetLineDash) else {
        panic!("a canvas answers its dash list");
    };
    assert_eq!(dash, vec![5.0, 5.0]);
}

#[test]
fn the_transform_reads_back() {
    let (_doc, target) = canvas_on_the_page("transform");
    paint(&target, Op2D::ResetTransform);
    paint(&target, Op2D::Scale(2.0, 3.0));
    let Query2DValue::Matrix(m) = query(&target, Query2D::GetTransform) else {
        panic!("a canvas answers its transform");
    };
    assert!((m[0] - 2.0).abs() < 1e-4, "a = {}", m[0]);
    assert!((m[3] - 3.0).abs() < 1e-4, "d = {}", m[3]);
}

#[test]
fn a_canvas_can_hand_the_page_its_own_pixels() {
    let (_doc, target) = canvas_on_the_page("export");
    paint(&target, Op2D::SetFillStyleCss("#0a0".into()));
    paint(&target, Op2D::FillRect(0.0, 0.0, 20.0, 20.0));
    let Query2DValue::Bytes(png) = query(
        &target,
        Query2D::ToBlob {
            mime: "image/png".into(),
            quality: None,
        },
    ) else {
        panic!("a canvas encodes to PNG");
    };
    assert_eq!(&png[..8], b"\x89PNG\r\n\x1a\n", "not a PNG signature");

    let url = text(query(
        &target,
        Query2D::ToDataUrl {
            mime: "image/png".into(),
            quality: None,
        },
    ));
    assert!(url.starts_with("data:image/png;base64,"), "got {}", &url[..30.min(url.len())]);
}

#[test]
fn an_absent_surface_answers_absent_and_never_a_default() {
    // The reason questions do not travel as ops. A caller has to be able to
    // tell "there is no canvas" from "the answer happens to be zero".
    install();
    assert!(matches!(
        query("n999999", Query2D::MeasureText("x".into())),
        Query2DValue::Absent
    ));
    assert!(matches!(
        query("n999999", Query2D::GetTransform),
        Query2DValue::Absent
    ));
    assert!(matches!(
        query(
            "n999999",
            Query2D::IsPointInPath {
                x: 0.0,
                y: 0.0,
                rule: "nonzero".into()
            }
        ),
        Query2DValue::Absent
    ));
}

#[test]
fn the_members_that_had_no_wire_format_now_cross() {
    // Shadows, filters, gradients, the eight text-style attributes, `arcTo`,
    // `roundRect` and the full `ellipse` — implemented by both engines and
    // reachable from nothing until the seam carried them.
    use vybe_platform_web::canvas_backend::{GradientDef, GradientKind, PatternDef};
    let (_doc, target) = canvas_on_the_page("new-ops");

    for op in [
        Op2D::SetShadowColor("rgba(0, 0, 0, 0.5)".into()),
        Op2D::SetShadowBlur(8.0),
        Op2D::SetShadowOffsetX(2.0),
        Op2D::SetShadowOffsetY(2.0),
        Op2D::SetFilter("blur(2px) grayscale(0.5)".into()),
        Op2D::SetGlobalCompositeOperation("multiply".into()),
        Op2D::SetImageSmoothingQuality("high".into()),
        Op2D::SetDirection("rtl".into()),
        Op2D::SetLetterSpacing("2px".into()),
        Op2D::SetWordSpacing("4px".into()),
        Op2D::SetFontKerning("none".into()),
        Op2D::SetFontStretch("condensed".into()),
        Op2D::SetFontVariantCaps("small-caps".into()),
        Op2D::SetTextRendering("geometricPrecision".into()),
        Op2D::SetLang("en".into()),
        Op2D::ArcTo(10.0, 0.0, 10.0, 10.0, 5.0),
        Op2D::RoundRect {
            x: 0.0,
            y: 0.0,
            w: 40.0,
            h: 20.0,
            radii: [4.0, 4.0, 4.0, 4.0],
        },
        Op2D::EllipseFull {
            x: 20.0,
            y: 10.0,
            rx: 8.0,
            ry: 4.0,
            rotation: 0.5,
            start: 0.0,
            end: 3.14,
            ccw: false,
        },
        Op2D::FillWithRule("evenodd".into()),
        Op2D::ClipWithRule("nonzero".into()),
        Op2D::FillTextMaxWidth("squeezed".into(), 0.0, 10.0, 20.0),
        Op2D::StrokeTextMaxWidth("squeezed".into(), 0.0, 20.0, 20.0),
        Op2D::SetFillGradient(GradientDef {
            kind: GradientKind::Linear {
                x0: 0.0,
                y0: 0.0,
                x1: 40.0,
                y1: 0.0,
            },
            stops: vec![(0.0, "red".into()), (1.0, "blue".into())],
        }),
        Op2D::SetStrokeGradient(GradientDef {
            kind: GradientKind::Radial {
                x0: 10.0,
                y0: 10.0,
                r0: 0.0,
                x1: 10.0,
                y1: 10.0,
                r1: 8.0,
            },
            stops: vec![(0.0, "#fff".into()), (1.0, "#000".into())],
        }),
        Op2D::SetFillPattern(PatternDef {
            pixels: vec![255u8; 2 * 2 * 4],
            width: 2,
            height: 2,
            repetition: "repeat".into(),
        }),
        Op2D::PutImageDataDirty {
            pixels: vec![200u8; 4 * 4 * 4],
            width: 4,
            height: 4,
            dx: 0.0,
            dy: 0.0,
            dirty_x: 1.0,
            dirty_y: 1.0,
            dirty_w: 2.0,
            dirty_h: 2.0,
        },
        Op2D::DrawFocusIfNeeded(true),
        Op2D::Reset,
    ] {
        paint(&target, op);
    }

    // Still answering afterwards, and back to its initial state — `Reset` was
    // the last op.
    assert_eq!(
        text(query(
            &target,
            Query2D::GetStringAttribute(StringAttribute::Direction)
        )),
        "inherit",
        "reset did not restore the initial direction"
    );
}

#[test]
fn a_gradient_paints_a_ramp_and_not_a_flat_colour() {
    // A gradient assigned to `fillStyle` used to be stringified, fail to parse
    // as CSS, and be IGNORED — so the fill came out in whatever colour was
    // already set. Silently: the assignment succeeded.
    use vybe_platform_web::canvas_backend::{GradientDef, GradientKind};
    let (_doc, target) = canvas_on_the_page("gradient-paints");

    paint(&target, Op2D::SetFillStyleCss("#000000".into()));
    paint(
        &target,
        Op2D::SetFillGradient(GradientDef {
            kind: GradientKind::Linear {
                x0: 0.0,
                y0: 0.0,
                x1: 32.0,
                y1: 0.0,
            },
            stops: vec![(0.0, "#ff0000".into()), (1.0, "#0000ff".into())],
        }),
    );
    paint(&target, Op2D::FillRect(0.0, 0.0, 32.0, 4.0));

    let Query2DValue::Pixels { data, .. } = query(
        &target,
        Query2D::GetImageData {
            sx: 0,
            sy: 0,
            sw: 32,
            sh: 1,
        },
    ) else {
        panic!("a canvas has pixels to hand back");
    };
    let at = |x: usize| (data[x * 4], data[x * 4 + 2]);
    let (left_r, left_b) = at(1);
    let (right_r, right_b) = at(30);
    assert!(
        left_r > right_r && right_b > left_b,
        "no ramp: left {:?}, right {:?} — the gradient was ignored and the \
         previous flat colour painted instead",
        at(1),
        at(30)
    );
}

// ── Path2D ──────────────────────────────────────────────────────────────────

use vybe_platform_web::canvas_backend::{PathDef, PathOp2D};

/// A 20×20 square at (10, 10), as a `Path2D` would carry it.
fn square() -> PathDef {
    PathDef {
        ops: vec![
            PathOp2D::Rect {
                x: 10.0,
                y: 10.0,
                w: 20.0,
                h: 20.0,
            },
        ],
    }
}

#[test]
fn a_prebuilt_path_can_be_filled() {
    let (_doc, target) = canvas_on_the_page("path-fill");
    paint(&target, Op2D::SetFillStyleCss("#ff0000".into()));
    paint(&target, Op2D::FillPath(square(), "nonzero".into()));

    let Query2DValue::Pixels { data, .. } = query(
        &target,
        Query2D::GetImageData {
            sx: 0,
            sy: 0,
            sw: 32,
            sh: 32,
        },
    ) else {
        panic!("a canvas has pixels to hand back");
    };
    let at = |x: usize, y: usize| {
        let i = (y * 32 + x) * 4;
        [data[i], data[i + 1], data[i + 2], data[i + 3]]
    };
    assert_eq!(at(20, 20), [255, 0, 0, 255], "inside the path");
    assert_eq!(at(2, 2), [0, 0, 0, 0], "outside it");
}

#[test]
fn filling_a_prebuilt_path_leaves_the_current_path_alone() {
    // **The reason `Path2D` exists.** `fill(path)` must not disturb whatever
    // the context is half-way through describing — if it did, a page could not
    // interleave the two, which is exactly what the overload is for.
    let (_doc, target) = canvas_on_the_page("path-isolation");

    // Start describing a shape on the context, then fill an unrelated path.
    paint(&target, Op2D::BeginPath);
    paint(&target, Op2D::Rect(40.0, 40.0, 10.0, 10.0));
    paint(&target, Op2D::FillPath(square(), "nonzero".into()));

    // The context's own path must still be the rect it was building.
    assert!(
        matches!(
            query(
                &target,
                Query2D::IsPointInPath {
                    x: 45.0,
                    y: 45.0,
                    rule: "nonzero".into()
                }
            ),
            Query2DValue::Bool(true)
        ),
        "filling a Path2D destroyed the context's current path"
    );
}

#[test]
fn a_prebuilt_path_can_be_hit_tested_without_touching_the_context() {
    let (_doc, target) = canvas_on_the_page("path-hit");
    paint(&target, Op2D::BeginPath);

    let inside = query(
        &target,
        Query2D::IsPointInPathOf {
            path: square(),
            x: 20.0,
            y: 20.0,
            rule: "nonzero".into(),
        },
    );
    let outside = query(
        &target,
        Query2D::IsPointInPathOf {
            path: square(),
            x: 2.0,
            y: 2.0,
            rule: "nonzero".into(),
        },
    );
    assert!(matches!(inside, Query2DValue::Bool(true)));
    assert!(matches!(outside, Query2DValue::Bool(false)));

    // And the context's own path is still empty — a point cannot be inside it.
    assert!(matches!(
        query(
            &target,
            Query2D::IsPointInPath {
                x: 20.0,
                y: 20.0,
                rule: "nonzero".into()
            }
        ),
        Query2DValue::Bool(false)
    ));
}

#[test]
fn every_segment_kind_crosses_the_seam() {
    // One arm per segment in the conversion, so a kind the seam learns to carry
    // and a backend forgets to convert is a compile error. This is the runtime
    // half: that none of them panics and all of them reach the engine.
    let (_doc, target) = canvas_on_the_page("path-segments");
    let all = PathDef {
        ops: vec![
            PathOp2D::MoveTo(0.0, 0.0),
            PathOp2D::LineTo(10.0, 0.0),
            PathOp2D::QuadraticCurveTo {
                cx: 15.0,
                cy: 5.0,
                x: 10.0,
                y: 10.0,
            },
            PathOp2D::BezierCurveTo {
                cx1: 8.0,
                cy1: 14.0,
                cx2: 4.0,
                cy2: 14.0,
                x: 0.0,
                y: 10.0,
            },
            PathOp2D::ArcTo {
                x1: 0.0,
                y1: 5.0,
                x2: 5.0,
                y2: 5.0,
                radius: 2.0,
            },
            PathOp2D::Rect {
                x: 20.0,
                y: 20.0,
                w: 5.0,
                h: 5.0,
            },
            PathOp2D::RoundRect {
                x: 30.0,
                y: 30.0,
                w: 10.0,
                h: 8.0,
                radii: [2.0, 2.0, 2.0, 2.0],
            },
            PathOp2D::Arc {
                x: 15.0,
                y: 15.0,
                r: 3.0,
                start: 0.0,
                end: 3.14,
                ccw: false,
            },
            PathOp2D::Ellipse {
                x: 25.0,
                y: 8.0,
                rx: 4.0,
                ry: 2.0,
                rotation: 0.3,
                start: 0.0,
                end: 6.28,
                ccw: false,
            },
            PathOp2D::ClosePath,
        ],
    };
    paint(&target, Op2D::SetFillStyleCss("#00ff00".into()));
    paint(&target, Op2D::FillPath(all.clone(), "evenodd".into()));
    paint(&target, Op2D::StrokePath(all.clone()));
    paint(&target, Op2D::ClipPath(all.clone(), "nonzero".into()));
    paint(&target, Op2D::AppendPath(all));

    // Something landed, and the canvas still answers.
    let Query2DValue::Pixels { data, .. } = query(
        &target,
        Query2D::GetImageData {
            sx: 0,
            sy: 0,
            sw: 40,
            sh: 40,
        },
    ) else {
        panic!("a canvas has pixels to hand back");
    };
    assert!(
        data.chunks_exact(4).any(|p| p[3] > 0),
        "a path made of every segment kind drew nothing at all"
    );
}

#[test]
fn appending_a_path_folds_it_into_the_current_one() {
    // `ctx.addPath(shape)` — the context's own path gains the shape's
    // segments, so a later `fill()` covers both.
    let (_doc, target) = canvas_on_the_page("path-append");
    paint(&target, Op2D::BeginPath);
    paint(&target, Op2D::AppendPath(square()));
    assert!(
        matches!(
            query(
                &target,
                Query2D::IsPointInPath {
                    x: 20.0,
                    y: 20.0,
                    rule: "nonzero".into()
                }
            ),
            Query2DValue::Bool(true)
        ),
        "the appended path never reached the context's current path"
    );
}
