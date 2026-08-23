//! `Canvas` — a first-class drawable widget.
//!
//! Drop a `Canvas` into a Form like any other control. It owns its own
//! BITMAP, the way HTML §4.12.5 says a `<canvas>` element does, and user code
//! draws into it through [`Canvas::with_canvas`]. The widget's
//! [`PanelWidget::render`] impl blits that bitmap onto the active
//! `tiny_skia::Pixmap` each frame.
//!
//! **It used to keep a `RecordingCanvas` and replay the commands every frame.**
//! That is a display list, not a canvas, and the difference is what a page can
//! ASK. `getImageData`, `toDataURL`, `toBlob` and `isPointInPath` are answered
//! from pixels; a list of commands nobody has replayed has none, so those
//! members could only ever return a default. Painting immediately into a bitmap
//! the element owns is also what makes a drawing and a `getImageData` of it
//! agree by construction, rather than by both being replayed the same way.
//!
//! ## Why a separate widget vs. just `paint_overlay` on every widget
//!
//! Every `PanelWidget` gets a default-empty [`PanelWidget::paint_overlay`]
//! hook so user paint code can layer on top of any control. The `Canvas`
//! widget is the bare/empty version where there's no chrome to overlay
//! — it's pure drawing surface. Use it when you want a blank rectangle
//! the user owns entirely. Use `paint_overlay` on `Button` / `Label` /
//! etc. when you want to add custom decoration on top of standard
//! widget chrome.
//!
//! Both paths route through the same [`canvas::Canvas`] trait — the
//! difference is just which surface is drawn into.
//!
//! ## Standalone usage
//!
//! ```ignore
//! use vybe_widgets::{Form, Canvas as CanvasWidget};
//! use vybe_widgets::canvas::{Canvas, Color};
//!
//! let mut form = Form::new("Drawing");
//! let mut c = CanvasWidget::new()
//!     .with_name("art")
//!     .with_background(Color::rgb(255, 255, 255));
//!
//! c.with_canvas(|canvas| {
//!     canvas.set_fill_color(Color::rgb(220, 50, 50));
//!     canvas.fill_rect(10.0, 10.0, 100.0, 100.0);
//! });
//!
//! form.add_control(c, 0.0, 0.0, 800.0, 600.0);
//! ```
//!
//! Zero VM. Zero host. Zero .NET wrapper.

use crate::canvas::{CanvasState, Color, TinySkiaCanvas};
use tiny_skia::Pixmap;
use crate::layout::{KeyEvent, LayoutRect, MouseEvent, PanelWidget, RenderContext, WidgetId};

/// A blank drawable surface widget.
///
/// Owns a `RecordingCanvas`. Replays it onto the active pixmap each
/// frame. Drawing persists across frames until the caller `clear()`s
/// the recording — matching .NET Graphics / VB6 / HTML5 canvas
/// semantics.
pub struct Canvas {
    name: String,
    rect: LayoutRect,
    id: WidgetId,
    background: Option<Color>,
    /// The element's bitmap — transparent black until something draws.
    ///
    /// Sized by `width`/`height`, NOT by the widget's layout rect: HTML
    /// §4.12.5 gives a canvas a bitmap of its own dimensions and then scales it
    /// into whatever box CSS gives it. That separation is why a page can size
    /// the bitmap for a device's pixel ratio while leaving the box alone.
    bitmap: Pixmap,
    /// Drawing state that survives between calls — see `canvas::CanvasState`.
    /// A page sets `fillStyle` in one call and fills in the next.
    state: CanvasState,
    /// Set once something assigns `width`/`height`. From then on the bitmap is
    /// the author's business and HTML §4.12.5 governs it exactly: the bitmap is
    /// the size they asked for, and the box scales it. Until then the widget
    /// sizes the backing store itself — see `ensure_backing`.
    explicit_size: bool,
    /// Device pixels per CSS pixel the bitmap is currently allocated for.
    device_scale: f32,
    /// Glyph raster cache for canvas text. Owned rather than borrowed because
    /// drawing happens when the PAGE calls, not while a frame is being painted,
    /// so the render context's cache is not in scope at the moment it is
    /// needed. The `FontSystem` itself is the shared one
    /// (`ide_text::with_font_system`), so faces still resolve identically to
    /// the rest of the toolkit.
    glyphs: cosmic_text::SwashCache,
}

/// A canvas with no `width`/`height` is 300 × 150 — HTML §4.12.5.
const DEFAULT_W: u32 = 300;
const DEFAULT_H: u32 = 150;

impl Canvas {
    /// Construct an empty canvas at the spec's default size.
    pub fn new() -> Self {
        Self {
            name: String::new(),
            rect: LayoutRect::zero(),
            id: WidgetId::next(),
            background: None,
            bitmap: Pixmap::new(DEFAULT_W, DEFAULT_H).expect("300x150 is a valid pixmap"),
            state: CanvasState::default(),
            explicit_size: false,
            device_scale: 1.0,
            glyphs: cosmic_text::SwashCache::new(),
        }
    }

    /// Set the widget's name (used for control lookup by the host
    /// bridge and the form's `send_command` API).
    pub fn with_name(mut self, name: &str) -> Self {
        self.name = name.to_string();
        self
    }

    /// Set the widget's background fill colour. `None` (the default)
    /// leaves the background transparent — whatever's underneath the
    /// widget shows through.
    pub fn with_background(mut self, color: Color) -> Self {
        self.background = Some(color);
        self
    }

    /// `canvas.width` / `canvas.height` — the BITMAP's size in pixels.
    ///
    /// Assigning either **reinitialises the bitmap to transparent black and
    /// resets the drawing state**, even when the value assigned is the one
    /// already there (HTML §4.12.5). `canvas.width = canvas.width` is the
    /// documented way a page clears a canvas, so preserving the pixels here
    /// would break it silently.
    pub fn set_bitmap_size(&mut self, width: u32, height: u32) {
        let (w, h) = (width.max(1), height.max(1));
        if let Some(fresh) = Pixmap::new(w, h) {
            self.bitmap = fresh;
            self.state = CanvasState::default();
            // An explicit size means the author is managing the backing store,
            // so the widget stops managing it — including the device scale.
            // A page that wants sharp pixels on a HiDPI display sets
            // `width = cssWidth * devicePixelRatio` itself, which is the same
            // thing `ensure_backing` does for a canvas that never said.
            self.explicit_size = true;
            self.device_scale = 1.0;
        }
    }

    /// Give the bitmap the size and resolution its box actually needs.
    ///
    /// **This is what keeps a canvas sharp, and what keeps it the right size.**
    /// The bitmap is allocated at `rect * scale` device pixels and the context
    /// is told the ratio, so drawing coordinates stay CSS pixels while the
    /// pixels behind them are real. Without the first half a canvas is an
    /// upscale of a 300x150 image on any display; without the second, every
    /// coordinate a caller passes means something different depending on the
    /// monitor.
    ///
    /// Existing content is RESAMPLED into the new bitmap rather than dropped. A
    /// resize is not an author's `canvas.width` assignment — nobody asked for a
    /// clear, and a window drag that erased the drawing would be a bug. The
    /// author's assignment goes through `set_bitmap_size`, which does clear.
    fn ensure_backing(&mut self, scale: f32) {
        if self.explicit_size || !scale.is_finite() || scale <= 0.0 {
            return;
        }
        if self.rect.w <= 0.0 || self.rect.h <= 0.0 {
            return;
        }
        let w = ((self.rect.w * scale).round() as u32).max(1);
        let h = ((self.rect.h * scale).round() as u32).max(1);
        if (w, h) == self.bitmap_size() && (scale - self.device_scale).abs() < f32::EPSILON {
            return;
        }
        let Some(mut fresh) = Pixmap::new(w, h) else { return };
        let (ow, oh) = self.bitmap_size();
        if ow > 0 && oh > 0 {
            fresh.draw_pixmap(
                0,
                0,
                self.bitmap.as_ref(),
                &tiny_skia::PixmapPaint::default(),
                tiny_skia::Transform::from_scale(w as f32 / ow as f32, h as f32 / oh as f32),
                None,
            );
        }
        self.bitmap = fresh;
        self.device_scale = scale;
        self.state.set_device_scale(scale);
    }

    /// The bitmap's dimensions — `canvas.width` / `canvas.height` read back.
    pub fn bitmap_size(&self) -> (u32, u32) {
        (self.bitmap.width(), self.bitmap.height())
    }

    /// The bitmap itself, for the pixel members: `getImageData`, `toDataURL`,
    /// `toBlob` and a `--capture`.
    pub fn bitmap(&self) -> &Pixmap {
        &self.bitmap
    }

    /// Draw on this canvas through the WHATWG 2D context.
    ///
    /// The state persists between calls, because a page sets `fillStyle` in one
    /// call and fills in the next; see `canvas::CanvasState`. The shared
    /// `FontSystem` is borrowed for the duration so canvas text resolves the
    /// same faces as every other string in the toolkit.
    pub fn with_canvas<R>(&mut self, f: impl FnOnce(&mut dyn crate::canvas::Canvas) -> R) -> R {
        // A caller can draw before the first frame is painted, so the backing
        // store is brought up to date HERE and not only in `render` — otherwise
        // the first drawing of every canvas would land in a stale bitmap and be
        // resampled into place.
        self.ensure_backing(self.device_scale);
        let saved = std::mem::take(&mut self.state);
        let bitmap = &mut self.bitmap;
        let glyphs = &mut self.glyphs;
        let (out, state) = crate::ide_text::with_font_system(|fonts| {
            let mut canvas = TinySkiaCanvas::resume(bitmap, saved, Some((fonts, glyphs)));
            let out = f(&mut canvas);
            (out, canvas.suspend())
        });
        self.state = state;
        out
    }

    /// Back to a transparent bitmap and a default drawing state — HTML
    /// §4.12.5's `reset()`.
    pub fn clear(&mut self) {
        let (w, h) = self.bitmap_size();
        self.set_bitmap_size(w, h);
    }

    /// Whether nothing has been drawn — every pixel still transparent.
    pub fn is_blank(&self) -> bool {
        self.bitmap.pixels().iter().all(|p| p.alpha() == 0)
    }
}

// Persistence semantics: a drawing stays until the caller `clear()`s it,
// because it is IN the bitmap. That is the same guarantee the recording used
// to give by being replayed every frame, reached the way the spec reaches it.

impl Default for Canvas {
    fn default() -> Self {
        Self::new()
    }
}

impl PanelWidget for Canvas {
    fn name(&self) -> &str {
        &self.name
    }
    fn widget_id(&self) -> WidgetId {
        self.id
    }
    fn rect(&self) -> LayoutRect {
        self.rect
    }
    fn set_rect(&mut self, rect: LayoutRect) {
        self.rect = rect;
        self.ensure_backing(self.device_scale);
    }

    fn render(&mut self, ctx: &mut RenderContext) {
        if crate::canvas::trace_enabled() {
            // What there is to say about a bitmap is how much of it is inked.
            // The old line counted recorded COMMANDS, which no longer exist —
            // and a command count never answered the question anyone traces a
            // canvas to ask, which is whether anything landed.
            let (w, h) = self.bitmap_size();
            let inked = self.bitmap.pixels().iter().filter(|p| p.alpha() > 0).count();
            eprintln!(
                "[canvas] RENDER widget={:?} bitmap={w}x{h} inked_px={inked}",
                self.name,
            );
        }
        let r = self.rect;
        if r.w <= 0.0 || r.h <= 0.0 {
            return;
        }

        // 1. Background fill (if requested).
        if let Some(bg) = self.background {
            let mut paint = tiny_skia::Paint::default();
            paint.set_color_rgba8(bg.r, bg.g, bg.b, bg.a);
            paint.anti_alias = true;
            if let Some(rect) = tiny_skia::Rect::from_xywh(
                r.x * ctx.scale,
                r.y * ctx.scale,
                r.w * ctx.scale,
                r.h * ctx.scale,
            ) {
                ctx.pixmap
                    .fill_rect(rect, &paint, tiny_skia::Transform::identity(), None);
            }
        }

        // 2. Blit the element's bitmap into its box.
        //
        // Everything the page drew is already in those pixels — it was painted
        // when the call was made, not queued for now — so rendering a canvas is
        // rendering its bitmap and nothing else. That is also what the spec
        // says a canvas is.
        //
        // The bitmap is scaled to the box rather than clipped to it, because
        // `width`/`height` size the BITMAP and CSS sizes the BOX, and §4.12.5
        // maps one onto the other. A page that wants pixel-for-pixel sets the
        // two to match — which is exactly what `devicePixelRatio` handling is.
        // The device scale is only knowable here, so this is where a change of
        // it (a window moved between a 1x and a 2x display) reaches the bitmap.
        self.ensure_backing(ctx.scale);
        let (bw, bh) = self.bitmap_size();
        if bw > 0 && bh > 0 && r.w > 0.0 && r.h > 0.0 {
            let sx = (r.w * ctx.scale) / bw as f32;
            let sy = (r.h * ctx.scale) / bh as f32;
            ctx.pixmap.draw_pixmap(
                0,
                0,
                self.bitmap.as_ref(),
                &tiny_skia::PixmapPaint::default(),
                tiny_skia::Transform::from_translate(r.x * ctx.scale, r.y * ctx.scale)
                    .pre_scale(sx, sy),
                None,
            );
        }
    }

    fn handle_mouse(&mut self, _event: &MouseEvent) -> bool {
        false
    }
    fn handle_key(&mut self, _event: &KeyEvent) -> bool {
        false
    }

    /// Used by the host bridge to find this widget by name and pull
    /// out its underlying `RecordingCanvas`. See
    /// `vybe_host::gui_state::GuiState::find_canvas_mut`.
    fn as_any_mut(&mut self) -> Option<&mut dyn std::any::Any> {
        Some(self as &mut dyn std::any::Any)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::canvas::{Canvas as CanvasTrait, Color};

    /// The pixel at (x, y), premultiplied RGBA.
    fn px(c: &Canvas, x: u32, y: u32) -> [u8; 4] {
        let p = c.bitmap().pixels()[(y * c.bitmap().width() + x) as usize];
        [p.red(), p.green(), p.blue(), p.alpha()]
    }

    #[test]
    fn a_new_canvas_is_the_specs_default_size_and_transparent() {
        let c = Canvas::new();
        assert_eq!(c.bitmap_size(), (300, 150));
        assert!(c.is_blank());
    }

    #[test]
    fn the_drawing_state_survives_between_calls() {
        // The reason a context is retained rather than rebuilt: a page sets
        // `fillStyle` in one call and fills in the next, and each call reaches
        // the engine on its own.
        let mut c = Canvas::new();
        c.with_canvas(|ctx| ctx.set_fill_color(Color::rgb(0, 0, 255)));
        c.with_canvas(|ctx| ctx.fill_rect(0.0, 0.0, 20.0, 20.0));
        assert_eq!(px(&c, 10, 10), [0, 0, 255, 255]);
    }

    #[test]
    fn what_was_drawn_can_be_read_back() {
        // **The whole point of behavioural parity.** A recording could not
        // answer this at all: at the moment the page asks, there are no pixels
        // to read. Now the drawing IS the pixels.
        let mut c = Canvas::new();
        c.with_canvas(|ctx| {
            ctx.set_fill_color(Color::rgb(255, 128, 0));
            ctx.fill_rect(5.0, 5.0, 10.0, 10.0);
        });
        let data = c
            .with_canvas(|ctx| ctx.get_image_data(0, 0, 20, 20))
            .expect("a canvas has pixels to hand back");
        // `ImageData` is STRAIGHT RGBA, not premultiplied — §4.12.5.
        let at = |x: usize, y: usize| {
            let i = (y * 20 + x) * 4;
            [
                data.data[i],
                data.data[i + 1],
                data.data[i + 2],
                data.data[i + 3],
            ]
        };
        assert_eq!(at(10, 10), [255, 128, 0, 255], "inside the rect");
        assert_eq!(at(1, 1), [0, 0, 0, 0], "outside it");
    }

    #[test]
    fn a_point_can_be_tested_against_the_path() {
        // `isPointInPath` needs the geometry, which a command list has and a
        // bitmap does not — so this proves the port kept the PATH state too,
        // not only the pixels.
        let mut c = Canvas::new();
        c.with_canvas(|ctx| {
            ctx.begin_path();
            ctx.rect(10.0, 10.0, 50.0, 50.0);
        });
        use crate::canvas::FillRule;
        assert!(c.with_canvas(|ctx| ctx.is_point_in_path(30.0, 30.0, FillRule::NonZero)));
        assert!(!c.with_canvas(|ctx| ctx.is_point_in_path(5.0, 5.0, FillRule::NonZero)));
    }

    #[test]
    fn the_canvas_can_hand_the_page_its_own_pixels() {
        // `toDataURL` / `toBlob` — a page cannot be given a file path, so this
        // is the only export route the spec gives a canvas.
        let mut c = Canvas::new();
        c.with_canvas(|ctx| {
            ctx.set_fill_color(Color::rgb(10, 200, 30));
            ctx.fill_rect(0.0, 0.0, 300.0, 150.0);
        });
        let png = c
            .with_canvas(|ctx| ctx.to_blob("image/png", None))
            .expect("a PNG");
        assert_eq!(&png[..8], b"\x89PNG\r\n\x1a\n", "not a PNG signature");

        let url = c.with_canvas(|ctx| ctx.to_data_url("image/png", None));
        assert!(
            url.starts_with("data:image/png;base64,"),
            "toDataURL answered {}",
            &url[..url.len().min(40)]
        );
    }

    #[test]
    fn canvas_text_actually_draws() {
        // Text is drawn when the PAGE calls, not while a frame is painted, so
        // the fonts have to be reachable at that moment. If they are not,
        // `fillText` returns having done nothing and every test that does not
        // sample glyph pixels still passes.
        let mut c = Canvas::new();
        c.with_canvas(|ctx| {
            ctx.set_fill_color(Color::rgb(0, 0, 0));
            ctx.set_font(&crate::canvas::Font::new("sans-serif", 48.0));
            ctx.fill_text("HHHH", 5.0, 60.0);
        });
        let inked = c.bitmap().pixels().iter().filter(|p| p.alpha() > 0).count();
        assert!(inked > 200, "fillText drew {inked} opaque pixels");
    }

    #[test]
    fn measuring_text_sees_the_font_in_effect() {
        let mut c = Canvas::new();
        c.with_canvas(|ctx| ctx.set_font(&crate::canvas::Font::new("sans-serif", 12.0)));
        let small = c.with_canvas(|ctx| ctx.measure_text("HHHHHHHH").width);
        c.with_canvas(|ctx| ctx.set_font(&crate::canvas::Font::new("sans-serif", 48.0)));
        let large = c.with_canvas(|ctx| ctx.measure_text("HHHHHHHH").width);
        assert!(small > 0.0, "measured nothing at 12px");
        assert!(large > small * 2.0, "12px {small}, 48px {large}");
    }

    #[test]
    fn sizing_the_bitmap_clears_it_even_to_the_same_value() {
        // HTML §4.12.5 — `canvas.width = canvas.width` is the documented way a
        // page clears a canvas, which only works because ANY set reinitialises.
        let mut c = Canvas::new();
        c.with_canvas(|ctx| {
            ctx.set_fill_color(Color::rgb(255, 0, 0));
            ctx.fill_rect(0.0, 0.0, 300.0, 150.0);
        });
        assert!(!c.is_blank());
        c.set_bitmap_size(300, 150);
        assert!(c.is_blank(), "the bitmap was not reinitialised");

        // The fill colour went with it — this paints the default black.
        c.with_canvas(|ctx| ctx.fill_rect(0.0, 0.0, 10.0, 10.0));
        assert_eq!(px(&c, 5, 5), [0, 0, 0, 255], "the state was not reset");
    }

    #[test]
    fn shadows_and_filters_reach_the_bitmap() {
        // Two of the 41 members widgets did not have before the port, and the
        // two that need a whole second layer rather than a paint setting.
        let mut c = Canvas::new();
        c.with_canvas(|ctx| {
            ctx.set_shadow(&crate::canvas::Shadow {
                color: Color::rgba(0, 0, 0, 255),
                blur: 8.0,
                offset_x: 0.0,
                offset_y: 0.0,
            });
            ctx.set_fill_color(Color::rgb(255, 0, 0));
            ctx.fill_rect(100.0, 50.0, 40.0, 40.0);
        });
        // A blurred shadow puts ink OUTSIDE the rect it was cast by.
        assert!(
            px(&c, 95, 70)[3] > 0,
            "no shadow fell outside the rectangle"
        );

        let mut f = Canvas::new();
        f.with_canvas(|ctx| {
            ctx.set_filter("grayscale(1)");
            ctx.set_fill_color(Color::rgb(255, 0, 0));
            ctx.fill_rect(0.0, 0.0, 50.0, 50.0);
        });
        let p = px(&f, 25, 25);
        assert_eq!(p[0], p[1], "grayscale left the channels unequal");
        assert_eq!(p[1], p[2], "grayscale left the channels unequal");
    }
}

/// HiDPI: a bitmap bigger than the coordinate space, and a coordinate space
/// that does not know about it.
///
/// The failure these guard against is not a crash — it is a canvas that looks
/// almost right. Drawing that lands at half size, a hit test that misses by a
/// factor of two, or a `getTransform` that reports a scale nobody applied.
#[cfg(test)]
mod hidpi_tests {
    use super::*;
    use crate::canvas::{Canvas as CanvasTrait, Color, FillRule, Matrix};
    use crate::layout::PanelWidget;

    /// A canvas laid out at `w x h` CSS pixels on a display with `scale`
    /// device pixels to the CSS pixel.
    fn laid_out(w: f32, h: f32, scale: f32) -> Canvas {
        let mut c = Canvas::new();
        c.set_rect(LayoutRect {
            x: 0.0,
            y: 0.0,
            w,
            h,
        });
        c.ensure_backing(scale);
        c
    }

    fn px(c: &Canvas, x: u32, y: u32) -> [u8; 4] {
        let p = c.bitmap().pixels()[(y * c.bitmap().width() + x) as usize];
        [p.red(), p.green(), p.blue(), p.alpha()]
    }

    #[test]
    fn the_backing_store_is_allocated_in_device_pixels() {
        // Without this the bitmap stays 300x150 and gets stretched into the
        // box — a canvas that is an upscale of a small image on every display.
        let c = laid_out(400.0, 200.0, 2.0);
        assert_eq!(c.bitmap_size(), (800, 400));
    }

    #[test]
    fn a_drawing_keeps_its_css_pixel_size_on_a_hidpi_surface() {
        // **The bug this exists for.** A caller asking for a 100 CSS-pixel
        // square must get 100 CSS pixels on every display. If the device scale
        // were not applied under the coordinates, the same call would cover
        // half the box on a 2x screen — visibly wrong, and wrong in a way that
        // only shows up on one class of machine.
        let mut c = laid_out(400.0, 200.0, 2.0);
        c.with_canvas(|ctx| {
            ctx.set_fill_color(Color::rgb(255, 0, 0));
            ctx.fill_rect(0.0, 0.0, 100.0, 100.0);
        });
        // 100 CSS px at 2x is 200 device px: inside at 199, outside at 201.
        assert_eq!(px(&c, 199, 199), [255, 0, 0, 255], "the square fell short");
        assert_eq!(px(&c, 201, 201), [0, 0, 0, 0], "the square overran");
    }

    #[test]
    fn a_hit_test_answers_in_css_pixels_too() {
        // A page hit-tests with the coordinates it drew with, and with the
        // coordinates a mouse event gives it — both CSS pixels. Inverting
        // through the device scale as well would halve the point and answer
        // about somewhere else.
        let mut c = laid_out(400.0, 200.0, 2.0);
        c.with_canvas(|ctx| {
            ctx.begin_path();
            ctx.rect(10.0, 10.0, 100.0, 100.0);
        });
        assert!(
            c.with_canvas(|ctx| ctx.is_point_in_path(60.0, 60.0, FillRule::NonZero)),
            "a point well inside the path missed"
        );
        assert!(
            !c.with_canvas(|ctx| ctx.is_point_in_path(150.0, 150.0, FillRule::NonZero)),
            "a point outside the path hit"
        );
        // The giveaway if the scale leaked in: 30,30 is outside the CSS-pixel
        // rect only when the point is NOT being halved. (Halved it would be
        // 15,15 — inside — so this is the assertion that catches it.)
        assert!(c.with_canvas(|ctx| ctx.is_point_in_path(30.0, 30.0, FillRule::NonZero)));
        assert!(!c.with_canvas(|ctx| ctx.is_point_in_path(5.0, 5.0, FillRule::NonZero)));
    }

    #[test]
    fn the_page_never_sees_the_device_scale() {
        // A page that has done nothing must read the identity, on any display.
        let mut c = laid_out(400.0, 200.0, 2.0);
        let m = c.with_canvas(|ctx| ctx.get_transform());
        let i = Matrix::default();
        assert!(
            (m.a - i.a).abs() < 1e-6 && (m.d - i.d).abs() < 1e-6,
            "getTransform leaked the device scale: {m:?}"
        );

        // And its own transform reads back as its own.
        c.with_canvas(|ctx| ctx.scale(3.0, 3.0));
        let m = c.with_canvas(|ctx| ctx.get_transform());
        assert!((m.a - 3.0).abs() < 1e-6, "expected the page's 3x, got {m:?}");
    }

    #[test]
    fn reset_transform_does_not_destroy_the_device_scale() {
        // The whole reason the scale sits UNDER the page's matrix. One
        // `resetTransform()` past it and every later drawing is half size.
        let mut c = laid_out(400.0, 200.0, 2.0);
        c.with_canvas(|ctx| {
            ctx.translate(50.0, 50.0);
            ctx.reset_transform();
            ctx.set_fill_color(Color::rgb(0, 0, 255));
            ctx.fill_rect(0.0, 0.0, 100.0, 100.0);
        });
        assert_eq!(px(&c, 199, 199), [0, 0, 255, 255], "reset lost the scale");
        assert_eq!(px(&c, 201, 201), [0, 0, 0, 0]);
    }

    #[test]
    fn reset_keeps_the_scale_because_the_surface_still_has_it() {
        // `reset()` clears the bitmap and the state. It does not shrink the
        // bitmap, so the ratio between the two is unchanged.
        let mut c = laid_out(400.0, 200.0, 2.0);
        c.with_canvas(|ctx| {
            ctx.reset();
            ctx.set_fill_color(Color::rgb(0, 255, 0));
            ctx.fill_rect(0.0, 0.0, 100.0, 100.0);
        });
        assert_eq!(px(&c, 199, 199), [0, 255, 0, 255]);
    }

    #[test]
    fn resizing_the_box_keeps_the_drawing() {
        // A window drag is not a `canvas.width` assignment, and erasing on one
        // would be a bug. (The author's assignment DOES clear — see
        // `sizing_the_bitmap_clears_it_even_to_the_same_value`.)
        let mut c = laid_out(200.0, 200.0, 1.0);
        c.with_canvas(|ctx| {
            ctx.set_fill_color(Color::rgb(255, 0, 0));
            ctx.fill_rect(0.0, 0.0, 200.0, 200.0);
        });
        c.set_rect(LayoutRect {
            x: 0.0,
            y: 0.0,
            w: 400.0,
            h: 400.0,
        });
        assert_eq!(c.bitmap_size(), (400, 400));
        assert!(!c.is_blank(), "the drawing was erased by a resize");
        assert_eq!(px(&c, 200, 200), [255, 0, 0, 255]);
    }

    #[test]
    fn moving_between_displays_reallocates_at_the_new_ratio() {
        let mut c = laid_out(400.0, 200.0, 1.0);
        assert_eq!(c.bitmap_size(), (400, 200));
        c.ensure_backing(2.0);
        assert_eq!(c.bitmap_size(), (800, 400));
        // And coordinates still mean CSS pixels afterwards.
        c.with_canvas(|ctx| {
            ctx.set_fill_color(Color::rgb(255, 0, 0));
            ctx.fill_rect(0.0, 0.0, 100.0, 100.0);
        });
        assert_eq!(px(&c, 199, 199), [255, 0, 0, 255]);
    }

    #[test]
    fn an_explicit_size_hands_the_backing_store_to_the_author() {
        // §4.12.5: once `width`/`height` are set, the bitmap is exactly that
        // many pixels and the box scales it. The widget stops second-guessing,
        // because a page that sizes its own canvas is usually doing the
        // devicePixelRatio arithmetic itself and would get it applied twice.
        let mut c = laid_out(400.0, 200.0, 2.0);
        assert_eq!(c.bitmap_size(), (800, 400));
        c.set_bitmap_size(300, 150);
        assert_eq!(c.bitmap_size(), (300, 150));
        c.set_rect(LayoutRect {
            x: 0.0,
            y: 0.0,
            w: 900.0,
            h: 450.0,
        });
        c.ensure_backing(2.0);
        assert_eq!(
            c.bitmap_size(),
            (300, 150),
            "an author-sized bitmap must not be resized by layout"
        );
        // And its coordinates are bitmap pixels, with no device scale under them.
        c.with_canvas(|ctx| {
            ctx.set_fill_color(Color::rgb(255, 0, 0));
            ctx.fill_rect(0.0, 0.0, 100.0, 100.0);
        });
        assert_eq!(px(&c, 99, 99), [255, 0, 0, 255]);
        assert_eq!(px(&c, 101, 101), [0, 0, 0, 0]);
    }

    #[test]
    fn a_nonsense_scale_is_ignored_rather_than_making_the_matrix_singular() {
        // Every hit test inverts the base matrix; a zero scale would make that
        // fail and silently answer `false` to everything.
        let mut c = laid_out(400.0, 200.0, 2.0);
        let before = c.bitmap_size();
        c.ensure_backing(0.0);
        c.ensure_backing(f32::NAN);
        assert_eq!(c.bitmap_size(), before);
        assert!(c.with_canvas(|ctx| {
            ctx.begin_path();
            ctx.rect(0.0, 0.0, 100.0, 100.0);
            ctx.is_point_in_path(50.0, 50.0, FillRule::NonZero)
        }));
    }
}
