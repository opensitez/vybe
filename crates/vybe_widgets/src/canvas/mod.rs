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
pub use types::{Color, Font, FontStyle, FontWeight, Image, LineCap, LineJoin};

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
        _ => false }
}

/// Turn canvas tracing on or off at runtime (the debugger's `trace canvas`).
pub fn set_trace_enabled(on: bool) {
    CANVAS_TRACE.store(
        if on { 2 } else { 1 },
        std::sync::atomic::Ordering::Relaxed,
    );
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

    /// Add an ellipse with centre `(x, y)` and radii `(rx, ry)` to the
    /// current path.
    fn ellipse(&mut self, x: f32, y: f32, rx: f32, ry: f32);

    // ─── Drawing ────────────────────────────────────────────────────────

    /// Fill the current path with the current fill colour.
    fn fill(&mut self);

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

    // ─── Clipping ───────────────────────────────────────────────────────

    /// Use the current path as the clip region for subsequent draw
    /// operations. The path is consumed (matches HTML5 `clip()`).
    /// `save` / `restore` push/pop the clip along with the rest of the
    /// paint state.
    fn clip(&mut self);

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
