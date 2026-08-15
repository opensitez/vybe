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
    FillRule, FilterQuality, LineCap as TsLineCap, LineJoin as TsLineJoin, Mask, Paint, Path,
    PathBuilder, Pixmap, PixmapPaint, PixmapRef, Stroke as TsStroke, Transform,
};

use cosmic_text::{
    Attrs, Buffer, Color as CosmicColor, Family, FontSystem, Metrics, Shaping, SwashCache,
};

use super::{Canvas, Color, Font, FontStyle, FontWeight, Image, LineCap, LineJoin};

/// `Canvas` impl that paints into a `tiny_skia::Pixmap`.
///
/// Holds a mutable borrow of the target pixmap for its lifetime, plus
/// the current paint state. Constructed each frame by the form's render
/// loop and dropped at end-of-frame.
///
/// Optionally borrows a `FontSystem + SwashCache` from the caller to
/// support text rendering. When no font system is provided (e.g. when
/// constructed inline by tests that don't need text), `fill_text` /
/// `stroke_text` are no-ops. The form's render loop always provides one
/// — `RenderContext` already carries them — so live rendering of
/// recordings with text just works.
pub struct TinySkiaCanvas<'a> {
    pixmap: &'a mut Pixmap,
    state: PaintState,
    state_stack: Vec<PaintFrame>,
    path: PathBuilder,
    text_ctx: Option<TextCtx<'a>>,
    /// Active clip mask (if `clip` was called). Combined with the
    /// current transform when issuing draw ops. None = no clipping.
    clip_mask: Option<Mask>,
}

/// One entry on the `save`/`restore` stack — paint state + a snapshot
/// of the active clip mask. Cloning the `Mask` is the only way to
/// preserve clipping across `save`/`restore`; tiny-skia masks are just
/// alpha buffers, so the clone is a `Vec<u8>` copy of pixmap-sized
/// bytes — fine for typical UI usage.
#[derive(Clone)]
struct PaintFrame {
    state: PaintState,
    clip_mask: Option<Mask>,
}

/// Borrowed cosmic-text resources used for text rendering. Shared with
/// the rest of the toolkit via `RenderContext::font_system /
/// swash_cache`. Optional so callers that don't need text don't have
/// to set up cosmic-text.
struct TextCtx<'a> {
    font_system: &'a mut FontSystem,
    swash_cache: &'a mut SwashCache,
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
    dash_intervals: Vec<f32>,
    dash_offset: f32,
    image_smoothing: bool,
    text_align: super::TextAlign,
    text_baseline: super::TextBaseline,
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
            dash_intervals: Vec::new(),
            dash_offset: 0.0,
            // HTML5 canvas defaults `imageSmoothingEnabled` to true.
            image_smoothing: true,
            // …and `textAlign` to `start`, `textBaseline` to `alphabetic`.
            text_align: super::TextAlign::default(),
            text_baseline: super::TextBaseline::default(),
        }
    }
}

impl<'a> TinySkiaCanvas<'a> {
    /// Wrap a pixmap as a canvas. The canvas borrows the pixmap until
    /// it's dropped. Text rendering is disabled (no font system).
    pub fn new(pixmap: &'a mut Pixmap) -> Self {
        Self {
            pixmap,
            state: PaintState::default(),
            state_stack: Vec::new(),
            path: PathBuilder::new(),
            text_ctx: None,
            clip_mask: None,
        }
    }

    /// Wrap a pixmap as a canvas with text rendering enabled. The
    /// `FontSystem` and `SwashCache` are borrowed for the lifetime of
    /// the canvas. The form's render loop constructs canvases this way
    /// using `RenderContext::font_system / swash_cache`.
    pub fn with_text(
        pixmap: &'a mut Pixmap,
        font_system: &'a mut FontSystem,
        swash_cache: &'a mut SwashCache,
    ) -> Self {
        Self {
            pixmap,
            state: PaintState::default(),
            state_stack: Vec::new(),
            path: PathBuilder::new(),
            text_ctx: Some(TextCtx {
                font_system,
                swash_cache,
            }),
            clip_mask: None,
        }
    }

    /// Measure the logical width and height of `text` in the current
    /// font. Returns `(0.0, 0.0)` if no text context is attached.
    ///
    /// **Implemented here and reachable from nowhere.** `measureText` is the
    /// one canvas operation that ASKS rather than paints, and the seam only
    /// carries `Op2D`, which is fire-and-forget — so there is no wire format
    /// for an answer to come back through. `CanvasBackend` needs a query
    /// method before `web:canvas` can expose it, and `.NET`'s
    /// `Graphics.MeasureString` is the caller waiting on it. Same shape
    /// `ellipse` had: the engine had done the work and no page could reach it.
    pub fn measure_text(&mut self, text: &str) -> (f32, f32) {
        let Some(tc) = self.text_ctx.as_mut() else {
            return (0.0, 0.0);
        };
        let size = self.state.font.size;
        let metrics = Metrics::new(size, size * 1.3);
        let mut buf = Buffer::new(tc.font_system, metrics);
        let attrs = build_attrs(&self.state.font);
        buf.set_text(tc.font_system, text, &attrs, Shaping::Advanced, None);
        buf.shape_until_scroll(tc.font_system, false);
        let mut max_w = 0.0f32;
        let mut total_h = 0.0f32;
        for run in buf.layout_runs() {
            max_w = max_w.max(run.line_w);
            total_h = total_h.max(run.line_y + run.line_height);
        }
        (max_w, total_h)
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
    /// colour, line width, caps, joins, miter limit, dash, and global alpha.
    fn stroke_paint(&self) -> (Paint<'static>, TsStroke) {
        let mut p = Paint::default();
        let c = apply_alpha(self.state.stroke, self.state.global_alpha);
        p.set_color_rgba8(c.r, c.g, c.b, c.a);
        p.anti_alias = true;
        let mut stroke = TsStroke {
            width: self.state.line_width,
            line_cap: line_cap_to_ts(self.state.line_cap),
            line_join: line_join_to_ts(self.state.line_join),
            miter_limit: self.state.miter_limit,
            ..TsStroke::default()
        };
        if !self.state.dash_intervals.is_empty() {
            stroke.dash = tiny_skia::StrokeDash::new(
                self.state.dash_intervals.clone(),
                self.state.dash_offset,
            );
        }
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

    fn set_fill_color(&mut self, color: Color) {
        self.state.fill = color;
    }
    fn set_stroke_color(&mut self, color: Color) {
        self.state.stroke = color;
    }
    fn set_line_width(&mut self, width: f32) {
        self.state.line_width = width.max(0.0);
    }
    fn set_line_cap(&mut self, cap: LineCap) {
        self.state.line_cap = cap;
    }
    fn set_line_join(&mut self, join: LineJoin) {
        self.state.line_join = join;
    }
    fn set_miter_limit(&mut self, limit: f32) {
        self.state.miter_limit = limit.max(1.0);
    }
    fn set_global_alpha(&mut self, alpha: f32) {
        self.state.global_alpha = alpha.clamp(0.0, 1.0);
    }

    fn set_image_smoothing(&mut self, enabled: bool) {
        self.state.image_smoothing = enabled;
    }
    fn set_text_align(&mut self, align: super::TextAlign) {
        self.state.text_align = align;
    }
    fn set_text_baseline(&mut self, baseline: super::TextBaseline) {
        self.state.text_baseline = baseline;
    }
    fn set_font(&mut self, font: &Font) {
        self.state.font = font.clone();
    }
    fn set_line_dash(&mut self, intervals: &[f32]) {
        self.state.dash_intervals = intervals.to_vec();
    }
    fn set_line_dash_offset(&mut self, offset: f32) {
        self.state.dash_offset = offset;
    }

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
        let total = if ccw {
            -(start - end).abs()
        } else {
            (end - start).abs()
        };
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
        self.path
            .cubic_to(x + kx, y - ry, x + rx, y - ky, x + rx, y);
        self.path
            .cubic_to(x + rx, y + ky, x + kx, y + ry, x, y + ry);
        self.path
            .cubic_to(x - kx, y + ry, x - rx, y + ky, x - rx, y);
        self.path
            .cubic_to(x - rx, y - ky, x - kx, y - ry, x, y - ry);
        self.path.close();
    }

    // ─── Drawing ────────────────────────────────────────────────────────

    fn fill(&mut self) {
        if let Some(path) = self.take_path() {
            let paint = self.fill_paint();
            let mask = self.clip_mask.as_ref();
            self.pixmap
                .fill_path(&path, &paint, FillRule::Winding, self.state.transform, mask);
        }
    }

    fn stroke(&mut self) {
        if let Some(path) = self.take_path() {
            let (paint, stroke) = self.stroke_paint();
            let mask = self.clip_mask.as_ref();
            self.pixmap
                .stroke_path(&path, &paint, &stroke, self.state.transform, mask);
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
            let mask = self.clip_mask.as_ref();
            self.pixmap
                .fill_path(&path, &paint, FillRule::Winding, self.state.transform, mask);
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
            let mask = self.clip_mask.as_ref();
            self.pixmap
                .stroke_path(&path, &paint, &stroke, self.state.transform, mask);
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
            let mask = self.clip_mask.as_ref();
            self.pixmap
                .fill_path(&path, &paint, FillRule::Winding, self.state.transform, mask);
        }
    }

    fn fill_text(&mut self, text: &str, x: f32, y: f32) {
        // Real text rendering through cosmic-text. Requires a font
        // context — without one, this is a no-op (used by tests that
        // construct a TinySkiaCanvas without text support).
        let Some(tc) = self.text_ctx.as_mut() else {
            return;
        };

        // Build a buffer for the text, shape it, and rasterise each
        // glyph through the swash cache. The pixel data lands on the
        // pixmap via the existing `ide_text::draw_buffer` machinery —
        // we replicate the relevant bits inline so this module
        // doesn't depend on `super::super::ide_text` (which would
        // create a circular module dependency).
        // create a circular module dependency).
        let scale = self.state.transform.sx;
        let size = self.state.font.size;
        let metrics = Metrics::new(size, size * 1.3).scale(scale);
        let mut buf = Buffer::new(tc.font_system, metrics);
        let attrs = build_attrs(&self.state.font);
        buf.set_text(tc.font_system, text, &attrs, Shaping::Advanced, None);
        buf.shape_until_scroll(tc.font_system, false);

        let fill = apply_alpha(self.state.fill, self.state.global_alpha);
        let cosmic_color = CosmicColor::rgba(fill.r, fill.g, fill.b, fill.a);

        // Apply transform to the coordinates
        let px = x * self.state.transform.sx + self.state.transform.tx;
        let py = y * self.state.transform.sy + self.state.transform.ty;

        // **`textAlign` and `textBaseline` — what `x` and `y` actually name.**
        //
        // cosmic-text positions glyphs from the buffer's top-left, so drawing
        // at `(px, py)` unmodified means `x` = left edge and `y` = TOP. That is
        // `textAlign: left` with `textBaseline: top`, and the canvas spec's
        // defaults are `start` and **`alphabetic`** — `y` is the BASELINE, with
        // the glyphs sitting above it. Every string was landing about one
        // ascent too low, which the old comment here acknowledged rather than
        // fixed.
        //
        // Both offsets need the laid-out line, so they are read from the
        // shaped buffer rather than estimated: `line_w` is the advance and
        // `line_y` is the baseline's distance from the top.
        let (line_w, baseline) = buf
            .layout_runs()
            .next()
            .map(|run| (run.line_w, run.line_y))
            .unwrap_or((0.0, 0.0));
        let line_height = metrics.line_height;
        let px = match self.state.text_align {
            // Logical `start`/`end` equal left/right in a left-to-right
            // context, which is the only direction this canvas lays out.
            super::TextAlign::Start | super::TextAlign::Left => px,
            super::TextAlign::End | super::TextAlign::Right => px - line_w,
            super::TextAlign::Center => px - line_w / 2.0,
        };
        let py = match self.state.text_baseline {
            // `hanging` is not the same line as `top` in a font with real
            // metrics, but both sit at the head of the em box and this shaping
            // path surfaces only one of them — stated rather than pretended.
            super::TextBaseline::Top | super::TextBaseline::Hanging => py,
            super::TextBaseline::Middle => py - line_height / 2.0,
            super::TextBaseline::Alphabetic => py - baseline,
            super::TextBaseline::Ideographic | super::TextBaseline::Bottom => py - line_height,
        };

        crate::ide_text::draw_buffer(
            self.pixmap,
            tc.font_system,
            tc.swash_cache,
            &buf,
            px,
            py,
            cosmic_color,
        );
    }

    fn stroke_text(&mut self, text: &str, x: f32, y: f32) {
        // tiny-skia doesn't have outline-stroked glyphs out of the box.
        // For now, fall back to filled text in the stroke colour. This
        // matches what most lightweight canvas libs do; a real
        // implementation would tessellate each glyph's outline and
        // stroke the resulting path.
        let saved_fill = self.state.fill;
        self.state.fill = self.state.stroke;
        self.fill_text(text, x, y);
        self.state.fill = saved_fill;
    }

    fn clip(&mut self) {
        if let Some(path) = self.take_path() {
            // Rasterise the path into a fresh Mask the size of the
            // pixmap, then store it as the active clip. Subsequent
            // draws use the mask as a coverage modulator.
            let w = self.pixmap.width();
            let h = self.pixmap.height();
            let mut mask = match Mask::new(w, h) {
                Some(m) => m,
                None => return,
            };
            mask.fill_path(&path, FillRule::Winding, true, self.state.transform);
            // If there's already a clip mask, intersect them. tiny-skia's
            // Mask::intersect_path is the natural primitive for this.
            if let Some(existing) = self.clip_mask.as_mut() {
                existing.intersect_path(&path, FillRule::Winding, true, self.state.transform);
            } else {
                self.clip_mask = Some(mask);
            }
        }
    }

    fn reset_clip(&mut self) {
        self.clip_mask = None;
    }

    fn draw_image(&mut self, img: &Image, x: f32, y: f32, w: f32, h: f32) {
        // Build a tiny_skia::PixmapRef from the image's RGBA buffer,
        // then draw_pixmap with a scale transform that maps the
        // image's natural dimensions to the requested rect.
        if let Some(src) = PixmapRef::from_bytes(&img.pixels, img.width, img.height) {
            let scale_x = w / img.width as f32;
            let scale_y = h / img.height as f32;
            let xform = self
                .state
                .transform
                .pre_translate(x, y)
                .pre_scale(scale_x, scale_y);
            let pp = PixmapPaint {
                opacity: self.state.global_alpha,
                blend_mode: tiny_skia::BlendMode::SourceOver,
                // Smoothing OFF means nearest-neighbour, which is what a
                // software renderer upscaled to the window needs: bilinear
                // turns Doom's 320x200 frame into a blur.
                quality: if self.state.image_smoothing {
                    FilterQuality::Bilinear
                } else {
                    FilterQuality::Nearest
                },
            };
            self.pixmap.draw_pixmap(0, 0, src, &pp, xform, None);
        }
    }

    /// A raw pixel write — no transform, no clip, no alpha, no blending.
    ///
    /// Every other method here composes with `self.state`. This one must not,
    /// and it is the easiest arm in the file to get wrong by copying its
    /// neighbour: routing it through the paint state would make a software
    /// renderer's frame land somewhere else, tinted, or clipped away.
    /// `BlendMode::Source` REPLACES rather than composites, which is what
    /// "these are the pixels" means when the source has alpha.
    fn put_image_data(&mut self, img: &Image, dx: f32, dy: f32) {
        let Some(src) = PixmapRef::from_bytes(&img.pixels, img.width, img.height) else {
            return;
        };
        let pp = PixmapPaint {
            opacity: 1.0,
            blend_mode: tiny_skia::BlendMode::Source,
            quality: FilterQuality::Nearest,
        };
        // `draw_pixmap` takes an INTEGER destination, which suits a raw write:
        // the spec's `dx`/`dy` are longs, so there is no sub-pixel case to
        // lose here.
        self.pixmap.draw_pixmap(
            dx as i32,
            dy as i32,
            src,
            &pp,
            tiny_skia::Transform::identity(),
            None,
        );
    }

    // ─── State stack ────────────────────────────────────────────────────

    fn save(&mut self) {
        self.state_stack.push(PaintFrame {
            state: self.state.clone(),
            clip_mask: self.clip_mask.clone(),
        });
    }

    fn restore(&mut self) {
        if let Some(prev) = self.state_stack.pop() {
            self.state = prev.state;
            self.clip_mask = prev.clip_mask;
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

    /// tiny-skia's path builder already knows where the path ends, so this is
    /// a read rather than state of our own to keep in step.
    fn current_point(&self) -> Option<(f32, f32)> {
        self.path.last_point().map(|p| (p.x, p.y))
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
        LineCap::Butt => TsLineCap::Butt,
        LineCap::Round => TsLineCap::Round,
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

/// Build a `cosmic_text::Attrs` from a canvas `Font`.
///
/// Maps our `FontWeight` / `FontStyle` to cosmic-text's equivalents and
/// resolves the family name. Family is owned by the `Font` so we can
/// borrow it for the lifetime of the returned `Attrs`. The `'static`
/// hack via `Family::Name(&str)` means the borrow lifetime tracks the
/// `Font`'s — fine because `Attrs` doesn't escape this function.
fn build_attrs<'f>(font: &'f Font) -> Attrs<'f> {
    let family = if font.family.is_empty() || font.family == "sans-serif" {
        Family::SansSerif
    } else if font.family == "monospace" {
        Family::Monospace
    } else if font.family == "serif" {
        Family::Serif
    } else {
        Family::Name(&font.family)
    };
    let weight = match font.weight {
        FontWeight::Normal => cosmic_text::Weight::NORMAL,
        FontWeight::Bold => cosmic_text::Weight::BOLD,
    };
    let style = match font.style {
        FontStyle::Normal => cosmic_text::Style::Normal,
        FontStyle::Italic => cosmic_text::Style::Italic,
    };
    Attrs::new().family(family).weight(weight).style(style)
}
