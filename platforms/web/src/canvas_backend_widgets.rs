//! The painter behind `web:canvas`, resolving through the DOCUMENT.
//!
//! `getContext(element, "2d")` binds a context to a NODE (HTML §4.12.5), and
//! this is the half that turns that node into pixels: `Document::canvas_mut`
//! hands out the `<canvas>` element's own recording surface. Node → surface is
//! what a browser does internally, below the seam, so an engine swap replaces
//! this file and nothing above it.
//!
//! **Why it exists.** The only backend before this one resolved through
//! `GuiState` — it walked `GuiState.form.controls` for a widget whose NAME
//! matched. A canvas made by `createElement` lives in the document's form,
//! which is a different instance, so every lookup missed and fell through to
//! `GuiState.overlay_canvases`, a side map keyed by string. That map is
//! composited by `gui_capture` and by nothing else: `Form::render_overlays` has
//! no other caller, so a `--capture` showed the drawing while a real window
//! showed an empty `<canvas>`. A screenshot that proves nothing about the
//! window is the worst shape a bug can take, and it is what a second tree buys.
//!
//! The target string stays a string on purpose: a backend is below the seam and
//! may key its surfaces however it likes. What had to be spec-shaped is the
//! guest-facing call, and that is `canvas.rs`'s business.

use std::sync::Arc;

use crate::canvas_backend::{self, CanvasBackend, Op2D};
// `Canvas` is the DRAWING trait (`save`, `fill_rect`, …) that `RecordingCanvas`
// implements. The canvas WIDGET is no longer named here: `get_context_2d` hands
// back the surface directly, so this file never sees the element's widget.
use vybe_widgets::canvas::{Canvas as _, Color, Font, FontStyle, FontWeight, LineCap, LineJoin};
use vybe_widgets::dom::{self, NodeId};

struct DocumentBackend;

fn color(r: u8, g: u8, b: u8, a: u8) -> Color {
    Color { r, g, b, a }
}

/// The node a target names.
///
/// Two forms, in the order a caller means them:
///
/// 1. `n<id>` — what an element-bound context carries. `getContext` derives it
///    from the node it was given, so this is the direct case and no search
///    happens at all.
/// 2. A control NAME — .NET `CreateGraphics` and Flutter's canvas bridge still
///    pass one, and `Document::element_by_control_name` resolves it the way a
///    caller means it (the `name` attribute, then `id`, then the internal
///    widget name). MIGRATION ONLY: a caller that made the element already
///    holds the handle.
fn node_of(document: &dom::Document, target: &str) -> Option<NodeId> {
    if let Some(rest) = target.strip_prefix('n') {
        if let Ok(id) = rest.parse::<NodeId>() {
            return Some(id);
        }
    }
    document.element_by_control_name(target)
}

/// Borrow the surface `target` names, in the ambient document.
///
/// The closure receives the 2D CONTEXT directly. It used to receive the canvas
/// element's widget and every caller then asked it for the surface — that hop
/// is how a canvas happens to be stored here, and `getContext` has no name for
/// it.
fn with_canvas<T>(
    target: &str,
    f: impl FnOnce(&mut vybe_widgets::canvas::RecordingCanvas) -> T,
) -> Option<T> {
    dom::with_document(crate::html::active_document(), |document| {
        let node = node_of(document, target)?;
        document.get_context_2d(node).map(f)
    })
    .flatten()
}

impl CanvasBackend for DocumentBackend {
    /// `getContext`'s side effect. A `<canvas>` element already owns its
    /// surface — the element IS the storage — so there is nothing to create;
    /// touching it here only proves the node resolves.
    fn ensure(&self, target: &str) {
        let _ = with_canvas(target, |_| ());
    }

    /// Resolved through the DOCUMENT like everything else here — the font in
    /// effect belongs to the element's own surface, so a measurement taken
    /// anywhere else would be in the wrong font.
    fn measure_text(&self, target: &str, text: &str) -> Option<f32> {
        dom::with_document(crate::html::active_document(), |document| {
            let node = node_of(document, target)?;
            document.measure_text(node, text)
        })
        .flatten()
    }

    fn clear_all(&self, target: &str) {
        let _ = with_canvas(target, |c| c.clear());
    }

    fn apply(&self, target: &str, op: Op2D) {
        let _ = with_canvas(target, |c| {
            match op {
                Op2D::Save => c.save(),
                Op2D::Restore => c.restore(),
                Op2D::SetFillStyle(r, g, b, a) => c.set_fill_color(color(r, g, b, a)),
                Op2D::SetStrokeStyle(r, g, b, a) => c.set_stroke_color(color(r, g, b, a)),
                Op2D::SetLineWidth(w) => c.set_line_width(w),
                Op2D::SetLineDash(d) => c.set_line_dash(&d),
                Op2D::SetLineCap(k) => c.set_line_cap(match k.as_str() {
                    "round" => LineCap::Round,
                    "square" => LineCap::Square,
                    _ => LineCap::Butt,
                }),
                Op2D::SetLineJoin(k) => c.set_line_join(match k.as_str() {
                    "round" => LineJoin::Round,
                    "bevel" => LineJoin::Bevel,
                    _ => LineJoin::Miter,
                }),
                Op2D::SetGlobalAlpha(a) => c.set_global_alpha(a),
                Op2D::SetImageSmoothing(on) => c.set_image_smoothing(on),
                Op2D::SetFont {
                    family,
                    size,
                    bold,
                    italic,
                } => c.set_font(&Font {
                    family,
                    size,
                    weight: if bold {
                        FontWeight::Bold
                    } else {
                        FontWeight::Normal
                    },
                    style: if italic {
                        FontStyle::Italic
                    } else {
                        FontStyle::Normal
                    },
                }),
                Op2D::Translate(x, y) => c.translate(x, y),
                Op2D::Scale(x, y) => c.scale(x, y),
                Op2D::Rotate(a) => c.rotate(a),
                Op2D::BeginPath => c.begin_path(),
                Op2D::ClosePath => c.close_path(),
                Op2D::MoveTo(x, y) => c.move_to(x, y),
                Op2D::LineTo(x, y) => c.line_to(x, y),
                Op2D::Arc(x, y, r, s, e, ccw) => c.arc(x, y, r, s, e, ccw),
                Op2D::BezierCurveTo(a, b, cc, d, e, f) => c.bezier_curve_to(a, b, cc, d, e, f),
                Op2D::QuadraticCurveTo(a, b, cc, d) => c.quadratic_curve_to(a, b, cc, d),
                Op2D::Rect(x, y, w, h) => c.rect(x, y, w, h),
                Op2D::Ellipse(x, y, rx, ry) => c.ellipse(x, y, rx, ry),
                // `setTransform` is `ResetTransform` then `Transform`, emitted
                // as two ops — composing and replacing are different verbs.
                Op2D::Transform(a, b, cc, d, e, f) => c.transform(a, b, cc, d, e, f),
                Op2D::ResetTransform => c.reset_transform(),
                Op2D::SetMiterLimit(limit) => c.set_miter_limit(limit),
                Op2D::SetLineDashOffset(offset) => c.set_line_dash_offset(offset),
                // An unrecognised keyword leaves the current value, which is
                // what a browser does with a bad assignment to either of these
                // — it is not an error and it is not a reset.
                Op2D::SetTextAlign(k) => {
                    if let Some(a) = vybe_widgets::canvas::TextAlign::parse(&k) {
                        c.set_text_align(a);
                    }
                }
                Op2D::SetTextBaseline(k) => {
                    if let Some(b) = vybe_widgets::canvas::TextBaseline::parse(&k) {
                        c.set_text_baseline(b);
                    }
                }
                Op2D::Fill => c.fill(),
                Op2D::Stroke => c.stroke(),
                Op2D::Clip => c.clip(),
                Op2D::FillRect(x, y, w, h) => c.fill_rect(x, y, w, h),
                Op2D::StrokeRect(x, y, w, h) => c.stroke_rect(x, y, w, h),
                Op2D::ClearRect(x, y, w, h) => c.clear_rect(x, y, w, h),
                Op2D::FillText(t, x, y) => c.fill_text(&t, x, y),
                Op2D::StrokeText(t, x, y) => c.stroke_text(&t, x, y),
                Op2D::PutImageData {
                    pixels,
                    width,
                    height,
                    dx,
                    dy,
                } => {
                    let img = vybe_widgets::canvas::Image::from_rgba(width, height, pixels);
                    c.put_image_data(&img, dx, dy);
                }
                Op2D::DrawImageRgba {
                    pixels,
                    width,
                    height,
                    dx,
                    dy,
                    dw,
                    dh,
                } => {
                    let img = vybe_widgets::canvas::Image::from_rgba(width, height, pixels);
                    c.draw_image(&img, dx, dy, dw, dh);
                }
                Op2D::DrawImagePaletted {
                    indices,
                    palette,
                    width,
                    height,
                    dx,
                    dy,
                    dw,
                    dh,
                } => {
                    // The palette arrives as RGB triples (SDL's shape); the
                    // engine wants packed 0xRRGGBB entries.
                    let packed: Vec<u32> = palette
                        .chunks(3)
                        .map(|c| {
                            ((*c.first().unwrap_or(&0) as u32) << 16)
                                | ((*c.get(1).unwrap_or(&0) as u32) << 8)
                                | (*c.get(2).unwrap_or(&0) as u32)
                        })
                        .collect();
                    let img = vybe_widgets::canvas::Image::from_paletted(
                        width, height, &indices, &packed,
                    );
                    c.draw_image(&img, dx, dy, dw, dh);
                }
            }
        });
    }
}

/// Install the document as the surface `web:canvas` paints into.
pub fn install() {
    canvas_backend::set_backend(Arc::new(DocumentBackend));
}
