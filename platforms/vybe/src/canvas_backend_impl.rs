//! The painting engine behind `web:canvas`.
//!
//! `platforms/web` declares `CanvasRenderingContext2D` and owns nothing that
//! draws; this installs `vybe_widgets` as the engine that does. Swapping in a
//! real browser engine means replacing this file — the API, the guests, and
//! every adapter above it stay put.
//!
//! There is no `vybe:gui` canvas surface any more: guests call `web:canvas`.

use std::sync::{Arc, Mutex};

use vybe_platform_web::canvas_backend::{self, CanvasBackend, Op2D};
use vybe_widgets::canvas::{Canvas, Color, Font, FontStyle, FontWeight, LineCap, LineJoin};

use crate::gui_state::GuiState;

struct WidgetsBackend {
    gui: Arc<Mutex<GuiState>> }

fn color(r: u8, g: u8, b: u8, a: u8) -> Color {
    Color { r, g, b, a }
}

impl CanvasBackend for WidgetsBackend {
    fn ensure(&self, target: &str) {
        let _ = self.gui.lock().unwrap().find_canvas_mut(target);
    }

    fn clear_all(&self, target: &str) {
        self.gui.lock().unwrap().find_canvas_for_draw(target).clear();
    }

    fn apply(&self, target: &str, op: Op2D) {
        let mut gui = self.gui.lock().unwrap();
        let c = gui.find_canvas_for_draw(target);
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
                _ => LineCap::Butt }),
            Op2D::SetLineJoin(k) => c.set_line_join(match k.as_str() {
                "round" => LineJoin::Round,
                "bevel" => LineJoin::Bevel,
                _ => LineJoin::Miter }),
            Op2D::SetGlobalAlpha(a) => c.set_global_alpha(a),
            Op2D::SetImageSmoothing(on) => c.set_image_smoothing(on),
            Op2D::SetFont {
                family,
                size,
                bold,
                italic } => c.set_font(&Font {
                family,
                size,
                weight: if bold { FontWeight::Bold } else { FontWeight::Normal },
                style: if italic { FontStyle::Italic } else { FontStyle::Normal } }),
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
            Op2D::Fill => c.fill(),
            Op2D::Stroke => c.stroke(),
            Op2D::Clip => c.clip(),
            Op2D::FillRect(x, y, w, h) => c.fill_rect(x, y, w, h),
            Op2D::StrokeRect(x, y, w, h) => c.stroke_rect(x, y, w, h),
            Op2D::ClearRect(x, y, w, h) => c.clear_rect(x, y, w, h),
            Op2D::FillText(t, x, y) => c.fill_text(&t, x, y),
            Op2D::StrokeText(t, x, y) => c.stroke_text(&t, x, y),
            Op2D::DrawImageRgba {
                pixels,
                width,
                height,
                dx,
                dy,
                dw,
                dh } => {
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
                dh } => {
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
    }
}

/// Install `vybe_widgets` as the `web:canvas` painter.
pub fn install(gui: Arc<Mutex<GuiState>>) {
    canvas_backend::set_backend(Arc::new(WidgetsBackend { gui }));
}
