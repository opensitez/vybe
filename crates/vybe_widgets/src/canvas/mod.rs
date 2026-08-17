//! Canvas — generic immediate-mode drawing surface for `vybe_widgets`.
//!
//! This module is the **toolkit-level drawing primitive**. It is
//! deliberately VM-agnostic: a Rust user pulling in `vybe_widgets` as a
//! standalone toolkit can build a `Canvas` widget, paint on it via the
//! [`Canvas`] trait, and ship — with no `vybe_host`, `vybe_runtime`, or
//! .NET wrapper layer involved.
//!
//! ## Trait shape
//!
//! [`Canvas`] is HTML5-canvas-shaped. Every operation has a 1:1
//! counterpart in tiny-skia, Cairo, GDI, web canvas, Flutter canvas,
//! etc. — this is the primitive set every drawing backend converges on.
//! Higher-level framework façades that need to expose a different API
//! shape live OUTSIDE this crate; this trait stays the canonical
//! generic surface and never leaks framework-specific concepts (no
//! `Pen`, no `Brush`, no `Graphics`).
//!
//! ## Two impls
//!
//! - **[`RecordingCanvas`]** captures every call as a [`DrawCmd`]. The
//!   data IS the source of truth. Tests inspect it directly; the live
//!   render path replays it onto another canvas backend each frame.
//!
//! - **[`TinySkiaCanvas`]** paints onto a `tiny_skia::Pixmap`. Used by
//!   the form's render loop to turn a recording into pixels.
//!
//! Why two impls? Because the same canvas API has to serve two needs:
//! the live render must produce pixels, and tests must verify that the
//! right calls were made. Recording captures the calls as data,
//! tiny-skia turns them into pixels, and `RecordingCanvas::replay` is
//! the bridge between them.
//!
//! ## Standalone usage
//!
//! ```ignore
//! use vybe_widgets::canvas::{Canvas, TinySkiaCanvas, Color};
//! use tiny_skia::Pixmap;
//!
//! let mut pixmap = Pixmap::new(800, 600).unwrap();
//! pixmap.fill(tiny_skia::Color::WHITE);
//!
//! let mut canvas = TinySkiaCanvas::new(&mut pixmap);
//! canvas.set_fill_color(Color::rgb(255, 0, 0));
//! canvas.fill_rect(10.0, 10.0, 100.0, 100.0);
//!
//! canvas.set_stroke_color(Color::rgb(0, 0, 0));
//! canvas.set_line_width(3.0);
//! canvas.begin_path();
//! canvas.move_to(0.0, 0.0);
//! canvas.line_to(800.0, 600.0);
//! canvas.stroke();
//!
//! pixmap.save_png("out.png").unwrap();
//! ```

mod recording;
mod tinyskia;
mod types;

pub use recording::{DrawCmd, RecordingCanvas};
pub use tinyskia::TinySkiaCanvas;
pub use types::{
    Color, ColorStop, FillRule, Font, FontStyle, FontWeight, Gradient, GradientKind, Image,
    LineCap, LineJoin, Paint, Pattern, Repetition, Shadow, TextAlign, TextBaseline,
};

/// Canvas draw-routing tracing: `0` unread, `1` off, `2` on.
///
/// Seeded from `VYBE_DBG_CANVAS` on first read so the environment variable keeps
/// working, then settable at runtime by the debugger's `trace canvas on`. Every
/// trace site reads through [`trace_enabled`], so the switch is global.
static CANVAS_TRACE: std::sync::atomic::AtomicU8 = std::sync::atomic::AtomicU8::new(0);

/// True when canvas draw-routing tracing is on. Zero-cost when off: one relaxed
/// atomic load, no allocation, no environment lookup after the first call.
pub fn trace_enabled() -> bool {
    use std::sync::atomic::Ordering;
    match CANVAS_TRACE.load(Ordering::Relaxed) {
        0 => {
            let on = std::env::var_os("VYBE_DBG_CANVAS").is_some();
            CANVAS_TRACE.store(if on { 2 } else { 1 }, Ordering::Relaxed);
            on
        }
        2 => true,
        _ => false,
    }
}

/// Turn canvas tracing on or off at runtime (the debugger's `trace canvas`).
pub fn set_trace_enabled(on: bool) {
    CANVAS_TRACE.store(if on { 2 } else { 1 }, std::sync::atomic::Ordering::Relaxed);
}

/// HTML5-canvas-shaped immediate-mode drawing API.
///
/// Implementations:
/// - [`RecordingCanvas`] — captures calls as [`DrawCmd`] data.
/// - [`TinySkiaCanvas`] — paints onto a `tiny_skia::Pixmap`.
///
/// All coordinates are in pixels. Angles are in radians (matches HTML5
/// canvas; .NET-shaped wrappers are responsible for degrees → radians
/// conversion).
///
/// The trait is `?Sized`-friendly via dyn dispatch — every helper that
/// accepts a canvas takes `&mut dyn Canvas` so framework wrappers don't
/// need to know which concrete impl is in play.
pub trait Canvas {
    // ─── Paint state ────────────────────────────────────────────────────

    /// Set the colour used by subsequent `fill*` operations.
    fn set_fill_color(&mut self, color: Color);

    /// Set the colour used by subsequent `stroke*` operations.
    fn set_stroke_color(&mut self, color: Color);

    /// Set the line width (in pixels) used by subsequent `stroke*`
    /// operations.
    fn set_line_width(&mut self, width: f32);

    /// Set how the ends of stroked lines are drawn.
    fn set_line_cap(&mut self, cap: LineCap);

    /// Set how stroked lines join at corners.
    fn set_line_join(&mut self, join: LineJoin);

    /// Set the miter limit for sharp `LineJoin::Miter` corners.
    fn set_miter_limit(&mut self, limit: f32);

    /// Set the global alpha multiplier (0.0 .. 1.0) applied to all
    /// subsequent paint operations.
    fn set_global_alpha(&mut self, alpha: f32);

    /// HTML5 canvas `imageSmoothingEnabled`. `true` (the default) filters
    /// scaled images bilinearly; `false` selects nearest-neighbour, which is
    /// what a software-rendered frame upscaled to the window needs — bilinear
    /// blurs it. Defaulted so existing `Canvas` impls need no change.
    fn set_image_smoothing(&mut self, _enabled: bool) {}

    /// Set the font used by `fill_text` / `stroke_text`.
    fn set_font(&mut self, font: &Font);

    /// Set the dash pattern used by subsequent `stroke*` operations.
    /// Each f32 is alternating on/off in pixel units, matching HTML5
    /// canvas's `setLineDash`. Empty slice (the default) means solid.
    fn set_line_dash(&mut self, intervals: &[f32]);

    /// Set the dash phase offset used by subsequent `stroke*` ops.
    /// Default is 0.
    fn set_line_dash_offset(&mut self, offset: f32);

    /// `fillStyle` — a colour, a gradient or a pattern.
    ///
    /// `set_fill_color` is the colour-only spelling and stays the common path;
    /// this is the full one the spec types as
    /// `DOMString | CanvasGradient | CanvasPattern`.
    ///
    /// Defaulted to the flat-colour fallback so an impl that predates gradients
    /// keeps compiling AND keeps painting: a gradient degrades to its first
    /// stop rather than to nothing. An impl that can build a shader overrides
    /// this; `TinySkiaCanvas` does.
    fn set_fill_paint(&mut self, paint: &Paint) {
        self.set_fill_color(paint.as_flat_color());
    }

    /// `strokeStyle` — the stroke half of [`Canvas::set_fill_paint`].
    fn set_stroke_paint(&mut self, paint: &Paint) {
        self.set_stroke_color(paint.as_flat_color());
    }

    /// `shadowColor` / `shadowBlur` / `shadowOffsetX` / `shadowOffsetY`.
    ///
    /// One setter rather than four, because the four are meaningless apart:
    /// the spec draws a shadow only when the colour is non-transparent AND at
    /// least one of blur/offset is non-zero, so a backend has to read all four
    /// to answer "is there a shadow". Splitting them into four trait methods
    /// would make every impl reassemble the same tuple.
    ///
    /// Defaulted to a no-op: a backend without shadow support paints the shape
    /// unshadowed, which is a visible, honest degradation.
    fn set_shadow(&mut self, _shadow: &Shadow) {}

    // ─── Path building ──────────────────────────────────────────────────

    /// Reset the current path. Subsequent `move_to`, `line_to`, `arc`,
    /// etc. build a new path; `fill` / `stroke` paint it.
    fn begin_path(&mut self);

    /// Close the current sub-path by drawing a line back to its start.
    fn close_path(&mut self);

    /// Move the current point without drawing.
    fn move_to(&mut self, x: f32, y: f32);

    /// Draw a straight line from the current point to `(x, y)`.
    fn line_to(&mut self, x: f32, y: f32);

    /// Add a quadratic Bézier curve from the current point through
    /// `(cx, cy)` to `(x, y)`.
    fn quadratic_curve_to(&mut self, cx: f32, cy: f32, x: f32, y: f32);

    /// Add a cubic Bézier curve from the current point through
    /// `(cx1, cy1)` and `(cx2, cy2)` to `(x, y)`.
    fn bezier_curve_to(&mut self, cx1: f32, cy1: f32, cx2: f32, cy2: f32, x: f32, y: f32);

    /// Add an arc centred at `(x, y)` with radius `r`, sweeping from
    /// `start` to `end` radians. `ccw = true` reverses the sweep.
    fn arc(&mut self, x: f32, y: f32, r: f32, start: f32, end: f32, ccw: bool);

    /// Add a rectangle to the current path.
    fn rect(&mut self, x: f32, y: f32, w: f32, h: f32);

    /// Add a full ellipse with centre `(x, y)` and radii `(rx, ry)` to the
    /// current path.
    ///
    /// This is the closed-ellipse shorthand. The spec's `ellipse()` also takes
    /// a rotation and a start/end angle pair — see
    /// [`Canvas::ellipse_arc`], which this delegates to.
    fn ellipse(&mut self, x: f32, y: f32, rx: f32, ry: f32);

    /// `ellipse(x, y, radiusX, radiusY, rotation, startAngle, endAngle, ccw)`.
    ///
    /// **Provided, not required.** An ellipse arc is the unit circle arc under
    /// a scale-then-rotate transform, so it is expressible with `arc` and the
    /// transform stack every canvas already has — which is the spec's own
    /// definition, not an approximation of it.
    ///
    /// The transform is saved and restored around the arc so the path is the
    /// only thing that changes. Note the arc is appended to the CURRENT path:
    /// no `begin_path` here, because `ellipse` is a path-building method and
    /// the caller may be composing a compound path.
    ///
    /// Without this, `ellipse` could only draw a full, axis-aligned ellipse —
    /// a rotated one, or an elliptical wedge, was unexpressible.
    fn ellipse_arc(
        &mut self,
        x: f32,
        y: f32,
        rx: f32,
        ry: f32,
        rotation: f32,
        start: f32,
        end: f32,
        ccw: bool,
    ) {
        // A degenerate radius collapses the ellipse to a point or a segment;
        // the scale below would be singular, so take the closed-form shorthand
        // out and leave the path untouched rather than emit a NaN transform.
        if rx <= 0.0 || ry <= 0.0 {
            return;
        }
        self.save();
        self.translate(x, y);
        self.rotate(rotation);
        self.scale(rx, ry);
        // Radius 1 in the scaled space IS `rx`/`ry` in user space. The centre
        // is the origin because `translate` already moved it there.
        self.arc(0.0, 0.0, 1.0, start, end, ccw);
        self.restore();
    }

    /// `textAlign` — which end of the text the `x` names. Default `start`.
    fn set_text_align(&mut self, _align: TextAlign) {}

    /// `textBaseline` — which line of the text the `y` names. Default
    /// `alphabetic`, so `y` is the BASELINE and the glyphs sit above it.
    fn set_text_baseline(&mut self, _baseline: TextBaseline) {}

    /// `arcTo(x1, y1, x2, y2, radius)` — the arc that fillets the corner
    /// between the current point, `(x1, y1)` and `(x2, y2)`.
    ///
    /// **Provided, not required**: the spec defines it in terms of a line and
    /// an arc, so composing those is the definition rather than an
    /// approximation of it, and every canvas gets it without an implementation.
    ///
    /// The degenerate cases are the spec's own and are not merely guards — a
    /// zero radius, or three collinear points, is defined to add a straight
    /// line to `(x1, y1)` and stop. Rounding a rectangle whose corner radius is
    /// zero goes through here on every corner.
    fn arc_to(&mut self, x1: f32, y1: f32, x2: f32, y2: f32, radius: f32) {
        let (x0, y0) = match self.current_point() {
            Some(point) => point,
            // With no subpath the spec says start one at (x1, y1).
            None => {
                self.move_to(x1, y1);
                return;
            }
        };
        let (a1, b1) = (x0 - x1, y0 - y1);
        let (a2, b2) = (x2 - x1, y2 - y1);
        let len1 = (a1 * a1 + b1 * b1).sqrt();
        let len2 = (a2 * a2 + b2 * b2).sqrt();
        // Collinear or coincident, or no radius: a straight line, per spec.
        let cross = a1 * b2 - b1 * a2;
        if radius <= 0.0 || len1 == 0.0 || len2 == 0.0 || cross.abs() < f32::EPSILON {
            self.line_to(x1, y1);
            return;
        }
        let (u1, v1) = (a1 / len1, b1 / len1);
        let (u2, v2) = (a2 / len2, b2 / len2);
        // Half the angle between the two legs decides how far back along each
        // the tangent points sit.
        let cos_half = ((1.0 + (u1 * u2 + v1 * v2)) / 2.0).max(0.0).sqrt();
        let sin_half = ((1.0 - (u1 * u2 + v1 * v2)) / 2.0).max(0.0).sqrt();
        if sin_half <= f32::EPSILON {
            self.line_to(x1, y1);
            return;
        }
        let tangent = radius * cos_half / sin_half;
        let (t1x, t1y) = (x1 + u1 * tangent, y1 + v1 * tangent);
        let (t2x, t2y) = (x1 + u2 * tangent, y1 + v2 * tangent);
        // The centre sits on the bisector, one radius from each leg.
        let (bx, by) = (u1 + u2, v1 + v2);
        let blen = (bx * bx + by * by).sqrt();
        if blen <= f32::EPSILON {
            self.line_to(x1, y1);
            return;
        }
        let centre_distance = radius / sin_half;
        let (cx, cy) = (x1 + bx / blen * centre_distance, y1 + by / blen * centre_distance);
        let start = (t1y - cy).atan2(t1x - cx);
        let end = (t2y - cy).atan2(t2x - cx);
        self.line_to(t1x, t1y);
        // The sweep follows the turn direction of the two legs.
        self.arc(cx, cy, radius, start, end, cross > 0.0);
    }

    /// `roundRect(x, y, w, h, radius)` — a rectangle with rounded corners.
    ///
    /// Provided for the same reason as [`Canvas::arc_to`]: the spec builds it
    /// out of lines and arcs, so composing them IS the implementation.
    ///
    /// A radius larger than half the shorter side is scaled down rather than
    /// allowed to overlap, which is what the spec requires and what stops a
    /// large radius from turning the shape inside out.
    fn round_rect(&mut self, x: f32, y: f32, w: f32, h: f32, radius: f32) {
        if w == 0.0 || h == 0.0 {
            return;
        }
        // Negative width or height is legal and means the rectangle extends the
        // other way; normalising keeps the corner maths in one direction.
        let (x, w) = if w < 0.0 { (x + w, -w) } else { (x, w) };
        let (y, h) = if h < 0.0 { (y + h, -h) } else { (y, h) };
        let r = radius.max(0.0).min(w / 2.0).min(h / 2.0);
        self.move_to(x + r, y);
        self.line_to(x + w - r, y);
        self.arc_to(x + w, y, x + w, y + r, r);
        self.line_to(x + w, y + h - r);
        self.arc_to(x + w, y + h, x + w - r, y + h, r);
        self.line_to(x + r, y + h);
        self.arc_to(x, y + h, x, y + h - r, r);
        self.line_to(x, y + r);
        self.arc_to(x, y, x + r, y, r);
        self.close_path();
    }

    /// The current point of the path being built, if there is one.
    ///
    /// `arcTo` needs it — the arc is defined relative to where the path
    /// already is — and it is the one piece of path state a composed default
    /// cannot derive for itself. A canvas that does not track it answers
    /// `None`, and `arcTo` then starts a subpath, which is what the spec says
    /// to do in that case anyway.
    fn current_point(&self) -> Option<(f32, f32)> {
        None
    }

    // ─── Drawing ────────────────────────────────────────────────────────

    /// Fill the current path with the current fill colour.
    fn fill(&mut self);

    /// `fill(fillRule)` — fill the current path under an explicit winding rule.
    ///
    /// **Provided**, defaulting to [`Canvas::fill`] so `nonzero` (the spec's
    /// default) costs no impl change. A backend that can select the rule
    /// overrides this; one that cannot fills `evenodd` as `nonzero`, which
    /// differs only for self-intersecting paths and is visible rather than
    /// silent.
    fn fill_with_rule(&mut self, _rule: FillRule) {
        self.fill();
    }

    /// Stroke the current path with the current stroke colour and line
    /// width.
    fn stroke(&mut self);

    /// Fill a rectangle directly (does not modify the current path).
    fn fill_rect(&mut self, x: f32, y: f32, w: f32, h: f32);

    /// Stroke a rectangle directly (does not modify the current path).
    fn stroke_rect(&mut self, x: f32, y: f32, w: f32, h: f32);

    /// Clear a rectangle to fully transparent.
    fn clear_rect(&mut self, x: f32, y: f32, w: f32, h: f32);

    /// Fill `text` at baseline `(x, y)` with the current fill colour
    /// and font.
    fn fill_text(&mut self, text: &str, x: f32, y: f32);

    /// Stroke `text` at baseline `(x, y)` with the current stroke
    /// colour and font.
    fn stroke_text(&mut self, text: &str, x: f32, y: f32);

    /// Draw an image scaled into the rectangle `(x, y, w, h)`.
    fn draw_image(&mut self, img: &Image, x: f32, y: f32, w: f32, h: f32);

    /// `drawImage(image, sx, sy, sw, sh, dx, dy, dw, dh)` — the nine-argument
    /// form, which draws a SOURCE RECTANGLE of the image.
    ///
    /// This is what a sprite sheet needs, and it is not reducible to the
    /// four-argument form: that one always takes the whole image.
    ///
    /// **Provided**, by cropping the source rectangle into a new [`Image`] and
    /// drawing that. Copying is honest here rather than clever — the crop is
    /// the operation, and a backend that can blit a sub-rectangle directly
    /// should override this.
    ///
    /// A source rectangle outside the image is clamped, and an empty one draws
    /// nothing, which is the spec's own handling.
    fn draw_image_rect(
        &mut self,
        img: &Image,
        sx: f32,
        sy: f32,
        sw: f32,
        sh: f32,
        dx: f32,
        dy: f32,
        dw: f32,
        dh: f32,
    ) {
        let Some(cropped) = img.crop(sx, sy, sw, sh) else {
            return;
        };
        self.draw_image(&cropped, dx, dy, dw, dh);
    }

    /// `putImageData(imagedata, dx, dy)` — write pixels **directly** into the
    /// bitmap at `(dx, dy)`, one for one.
    ///
    /// Not a variant of [`Self::draw_image`], and the difference is the whole
    /// point of the operation: HTML §4.12.5 says `putImageData` is **not**
    /// affected by the current transform, the clipping region, `globalAlpha`
    /// or the compositing mode. It is a raw write, and that is why a software
    /// renderer uses it — the frame it computed is the frame that appears.
    ///
    /// It also does not scale: there is no `dw`/`dh`. Scaling raw pixels is
    /// done by putting them into a canvas and drawing THAT canvas, or by
    /// giving the `<canvas>` a CSS box larger than its bitmap. Both are the
    /// browser's answer; neither is a parameter here.
    fn put_image_data(&mut self, img: &Image, dx: f32, dy: f32);

    // ─── Clipping ───────────────────────────────────────────────────────

    /// Use the current path as the clip region for subsequent draw
    /// operations. The path is consumed (matches HTML5 `clip()`).
    /// `save` / `restore` push/pop the clip along with the rest of the
    /// paint state.
    fn clip(&mut self);

    /// `clip(fillRule)` — the clip half of [`Canvas::fill_with_rule`].
    fn clip_with_rule(&mut self, _rule: FillRule) {
        self.clip();
    }

    /// Reset the clip to the entire canvas.
    fn reset_clip(&mut self);

    // ─── State stack ────────────────────────────────────────────────────

    /// Push the current paint state (colours, line width, transform,
    /// etc.) onto an internal stack.
    fn save(&mut self);

    /// Pop the most recently saved paint state.
    fn restore(&mut self);

    // ─── Transforms ─────────────────────────────────────────────────────

    /// Translate the current transform by `(x, y)`.
    fn translate(&mut self, x: f32, y: f32);

    /// Rotate the current transform by `rad` radians.
    fn rotate(&mut self, rad: f32);

    /// Scale the current transform by `(sx, sy)`.
    fn scale(&mut self, sx: f32, sy: f32);

    /// Multiply the current transform by an arbitrary affine matrix.
    fn transform(&mut self, m11: f32, m12: f32, m21: f32, m22: f32, dx: f32, dy: f32);

    /// Reset the current transform to the identity.
    fn reset_transform(&mut self);
}
