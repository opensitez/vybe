//! `TinySkiaCanvas` — `Canvas` impl that paints onto a `tiny_skia::Pixmap`.
//!
//! This is the live render backend. The form's render loop wraps the
//! current frame's pixmap into a `TinySkiaCanvas` and either lets user
//! paint code call methods directly OR replays a `RecordingCanvas` onto
//! it (the standard path — recording is the persistent state).
//!
//! ## State management
//!
//! Canvas paint state (fill colour, stroke colour, line width, transform,
//! etc.) is non-trivial because tiny-skia doesn't carry it across calls
//! the way HTML5 canvas does. We track it explicitly in a
//! [`PaintState`] struct, mutated by the trait's `set_*` methods and
//! consulted by every `fill` / `stroke` call.
//!
//! `save` / `restore` push/pop the state stack. The trait's transform
//! ops compose into a `tiny_skia::Transform` carried in the current
//! state.
//!
//! Path building uses a `tiny_skia::PathBuilder` accumulated between
//! `begin_path` and `fill` / `stroke` calls. After painting, the
//! builder is reset for the next sub-path.
//!
//! ## Text rendering
//!
//! Text is rendered through `cosmic_text` (the same engine the rest of
//! the toolkit uses). The font cache is shared across calls via a
//! `&mut FontSystem` borrowed from the caller's
//! [`super::TextContext`]. For now, `fill_text` / `stroke_text` route
//! through `super::super::ide_text::draw_text` which is the existing
//! cosmic-text wrapper. (Stroke-text is approximated as fill-text
//! pending a real outline-stroke implementation.)
//!
//! ## Image rendering
//!
//! `draw_image` builds a `tiny_skia::Pixmap` from the image's RGBA
//! buffer and blits it onto the target with the appropriate scale
//! transform.

use tiny_skia::{
    Paint, FillRule, Stroke as TsStroke, LineCap as TsLineCap, LineJoin as TsLineJoin,
    Transform, PathBuilder, Path, Pixmap, PixmapPaint, FilterQuality,
    PixmapRef,
};

use super::{Canvas, Color, LineCap, LineJoin, Font, Image};

/// `Canvas` impl that paints into a `tiny_skia::Pixmap`.
///
/// Holds a mutable borrow of the target pixmap for its lifetime, plus
/// the current paint state. Constructed each frame by the form's render
/// loop and dropped at end-of-frame.
pub struct TinySkiaCanvas<'a> {
    pixmap: &'a mut Pixmap,
    state: PaintState,
    state_stack: Vec<PaintState>,
    path: PathBuilder,
}

#[derive(Clone)]
struct PaintState {
    fill: Color,
    stroke: Color,
    line_width: f32,
    line_cap: LineCap,
    line_join: LineJoin,
    miter_limit: f32,
    global_alpha: f32,
    font: Font,
    transform: Transform,
}

impl Default for PaintState {
    fn default() -> Self {
        Self {
            fill: Color::BLACK,
            stroke: Color::BLACK,
            line_width: 1.0,
            line_cap: LineCap::Butt,
            line_join: LineJoin::Miter,
            miter_limit: 10.0,
            global_alpha: 1.0,
            font: Font::default(),
            transform: Transform::identity(),
        }
    }
}

impl<'a> TinySkiaCanvas<'a> {
    /// Wrap a pixmap as a canvas. The canvas borrows the pixmap until
    /// it's dropped.
    pub fn new(pixmap: &'a mut Pixmap) -> Self {
        Self {
            pixmap,
            state: PaintState::default(),
            state_stack: Vec::new(),
            path: PathBuilder::new(),
        }
    }

    /// Build a `tiny_skia::Paint` from the current fill colour and
    /// global alpha.
    fn fill_paint(&self) -> Paint<'static> {
        let mut p = Paint::default();
        let c = apply_alpha(self.state.fill, self.state.global_alpha);
        p.set_color_rgba8(c.r, c.g, c.b, c.a);
        p.anti_alias = true;
        p
    }

    /// Build a `tiny_skia::Paint` + `Stroke` from the current stroke
    /// colour, line width, caps, joins, and global alpha.
    fn stroke_paint(&self) -> (Paint<'static>, TsStroke) {
        let mut p = Paint::default();
        let c = apply_alpha(self.state.stroke, self.state.global_alpha);
        p.set_color_rgba8(c.r, c.g, c.b, c.a);
        p.anti_alias = true;
        let stroke = TsStroke {
            width: self.state.line_width,
            line_cap: line_cap_to_ts(self.state.line_cap),
            line_join: line_join_to_ts(self.state.line_join),
            miter_limit: self.state.miter_limit,
            ..TsStroke::default()
        };
        (p, stroke)
    }

    /// Take the current path builder and replace it with an empty one.
    /// Used by `fill` / `stroke` after they've consumed the path.
    fn take_path(&mut self) -> Option<Path> {
        let pb = std::mem::replace(&mut self.path, PathBuilder::new());
        pb.finish()
    }
}

impl<'a> Canvas for TinySkiaCanvas<'a> {
    // ─── Paint state ────────────────────────────────────────────────────

    fn set_fill_color(&mut self, color: Color) { self.state.fill = color; }
    fn set_stroke_color(&mut self, color: Color) { self.state.stroke = color; }
    fn set_line_width(&mut self, width: f32) { self.state.line_width = width.max(0.0); }
    fn set_line_cap(&mut self, cap: LineCap) { self.state.line_cap = cap; }
    fn set_line_join(&mut self, join: LineJoin) { self.state.line_join = join; }
    fn set_miter_limit(&mut self, limit: f32) { self.state.miter_limit = limit.max(1.0); }
    fn set_global_alpha(&mut self, alpha: f32) { self.state.global_alpha = alpha.clamp(0.0, 1.0); }
    fn set_font(&mut self, font: &Font) { self.state.font = font.clone(); }

    // ─── Path building ──────────────────────────────────────────────────

    fn begin_path(&mut self) {
        self.path = PathBuilder::new();
    }

    fn close_path(&mut self) {
        self.path.close();
    }

    fn move_to(&mut self, x: f32, y: f32) {
        self.path.move_to(x, y);
    }

    fn line_to(&mut self, x: f32, y: f32) {
        self.path.line_to(x, y);
    }

    fn quadratic_curve_to(&mut self, cx: f32, cy: f32, x: f32, y: f32) {
        self.path.quad_to(cx, cy, x, y);
    }

    fn bezier_curve_to(&mut self, cx1: f32, cy1: f32, cx2: f32, cy2: f32, x: f32, y: f32) {
        self.path.cubic_to(cx1, cy1, cx2, cy2, x, y);
    }

    fn arc(&mut self, x: f32, y: f32, r: f32, start: f32, end: f32, ccw: bool) {
        // Polyline approximation. tiny-skia doesn't have a native arc
        // primitive — we sample N segments along the angular range. 32
        // segments is enough for visual smoothness up to ~100px radius.
        let segments = 32usize;
        let mut a = start;
        let total = if ccw { -(start - end).abs() } else { (end - start).abs() };
        let step = total / segments as f32;

        // First point — move_to so we don't connect to whatever was last.
        let (sx, sy) = (x + r * a.cos(), y + r * a.sin());
        if self.path.is_empty() {
            self.path.move_to(sx, sy);
        } else {
            self.path.line_to(sx, sy);
        }
        for _ in 0..segments {
            a += step;
            let (px, py) = (x + r * a.cos(), y + r * a.sin());
            self.path.line_to(px, py);
        }
    }

    fn rect(&mut self, x: f32, y: f32, w: f32, h: f32) {
        self.path.move_to(x, y);
        self.path.line_to(x + w, y);
        self.path.line_to(x + w, y + h);
        self.path.line_to(x, y + h);
        self.path.close();
    }

    fn ellipse(&mut self, x: f32, y: f32, rx: f32, ry: f32) {
        // Cubic-bezier approximation of an ellipse — same magic
        // constant trick as super::super::circle_path.
        let kx = rx * 0.5522848;
        let ky = ry * 0.5522848;
        self.path.move_to(x, y - ry);
        self.path.cubic_to(x + kx, y - ry, x + rx, y - ky, x + rx, y);
        self.path.cubic_to(x + rx, y + ky, x + kx, y + ry, x, y + ry);
        self.path.cubic_to(x - kx, y + ry, x - rx, y + ky, x - rx, y);
        self.path.cubic_to(x - rx, y - ky, x - kx, y - ry, x, y - ry);
        self.path.close();
    }

    // ─── Drawing ────────────────────────────────────────────────────────

    fn fill(&mut self) {
        if let Some(path) = self.take_path() {
            let paint = self.fill_paint();
            self.pixmap.fill_path(&path, &paint, FillRule::Winding, self.state.transform, None);
        }
    }

    fn stroke(&mut self) {
        if let Some(path) = self.take_path() {
            let (paint, stroke) = self.stroke_paint();
            self.pixmap.stroke_path(&path, &paint, &stroke, self.state.transform, None);
        }
    }

    fn fill_rect(&mut self, x: f32, y: f32, w: f32, h: f32) {
        let mut pb = PathBuilder::new();
        pb.move_to(x, y);
        pb.line_to(x + w, y);
        pb.line_to(x + w, y + h);
        pb.line_to(x, y + h);
        pb.close();
        if let Some(path) = pb.finish() {
            let paint = self.fill_paint();
            self.pixmap.fill_path(&path, &paint, FillRule::Winding, self.state.transform, None);
        }
    }

    fn stroke_rect(&mut self, x: f32, y: f32, w: f32, h: f32) {
        let mut pb = PathBuilder::new();
        pb.move_to(x, y);
        pb.line_to(x + w, y);
        pb.line_to(x + w, y + h);
        pb.line_to(x, y + h);
        pb.close();
        if let Some(path) = pb.finish() {
            let (paint, stroke) = self.stroke_paint();
            self.pixmap.stroke_path(&path, &paint, &stroke, self.state.transform, None);
        }
    }

    fn clear_rect(&mut self, x: f32, y: f32, w: f32, h: f32) {
        // Fill with transparent black, replacing whatever's there.
        let mut pb = PathBuilder::new();
        pb.move_to(x, y);
        pb.line_to(x + w, y);
        pb.line_to(x + w, y + h);
        pb.line_to(x, y + h);
        pb.close();
        if let Some(path) = pb.finish() {
            let mut paint = Paint::default();
            paint.set_color_rgba8(0, 0, 0, 0);
            paint.blend_mode = tiny_skia::BlendMode::Source;
            self.pixmap.fill_path(&path, &paint, FillRule::Winding, self.state.transform, None);
        }
    }

    fn fill_text(&mut self, _text: &str, _x: f32, _y: f32) {
        // Text rendering is intentionally a no-op at the canvas-trait
        // level — cosmic-text wants a `&mut FontSystem` borrowed from
        // the application's `TextContext`, which we don't have access
        // to from this signature. The form's render loop is responsible
        // for replaying text commands through its own text-renderer
        // path. (Recording captures `FillText` cmds; the runner
        // pattern-matches and dispatches them through cosmic-text.)
        //
        // For pure tinyskia-only usage (no cosmic-text dependency in
        // the toolkit consumer's app), this no-op is the safe default.
    }

    fn stroke_text(&mut self, _text: &str, _x: f32, _y: f32) {
        // Same reasoning as `fill_text` — handled by the runner.
    }

    fn draw_image(&mut self, img: &Image, x: f32, y: f32, w: f32, h: f32) {
        // Build a tiny_skia::PixmapRef from the image's RGBA buffer,
        // then draw_pixmap with a scale transform that maps the
        // image's natural dimensions to the requested rect.
        if let Some(src) = PixmapRef::from_bytes(&img.pixels, img.width, img.height) {
            let scale_x = w / img.width as f32;
            let scale_y = h / img.height as f32;
            let xform = self.state.transform
                .pre_translate(x, y)
                .pre_scale(scale_x, scale_y);
            let pp = PixmapPaint {
                opacity: self.state.global_alpha,
                blend_mode: tiny_skia::BlendMode::SourceOver,
                quality: FilterQuality::Bilinear,
            };
            self.pixmap.draw_pixmap(0, 0, src, &pp, xform, None);
        }
    }

    // ─── State stack ────────────────────────────────────────────────────

    fn save(&mut self) {
        self.state_stack.push(self.state.clone());
    }

    fn restore(&mut self) {
        if let Some(prev) = self.state_stack.pop() {
            self.state = prev;
        }
    }

    // ─── Transforms ─────────────────────────────────────────────────────

    fn translate(&mut self, x: f32, y: f32) {
        self.state.transform = self.state.transform.pre_translate(x, y);
    }

    fn rotate(&mut self, rad: f32) {
        self.state.transform = self.state.transform.pre_rotate(rad.to_degrees());
    }

    fn scale(&mut self, sx: f32, sy: f32) {
        self.state.transform = self.state.transform.pre_scale(sx, sy);
    }

    fn transform(&mut self, m11: f32, m12: f32, m21: f32, m22: f32, dx: f32, dy: f32) {
        let t = Transform::from_row(m11, m12, m21, m22, dx, dy);
        self.state.transform = self.state.transform.pre_concat(t);
    }

    fn reset_transform(&mut self) {
        self.state.transform = Transform::identity();
    }
}

// ─── Helpers ──────────────────────────────────────────────────────────────

/// Multiply a colour's alpha channel by `alpha` (0.0..1.0). Used by
/// `fill_paint` / `stroke_paint` so the canvas-trait `set_global_alpha`
/// works on top of any explicit colour alpha.
fn apply_alpha(color: Color, alpha: f32) -> Color {
    let a = (color.a as f32 * alpha.clamp(0.0, 1.0)).round() as u8;
    Color { a, ..color }
}

fn line_cap_to_ts(cap: LineCap) -> TsLineCap {
    match cap {
        LineCap::Butt   => TsLineCap::Butt,
        LineCap::Round  => TsLineCap::Round,
        LineCap::Square => TsLineCap::Square,
    }
}

fn line_join_to_ts(join: LineJoin) -> TsLineJoin {
    match join {
        LineJoin::Miter => TsLineJoin::Miter,
        LineJoin::Round => TsLineJoin::Round,
        LineJoin::Bevel => TsLineJoin::Bevel,
    }
}
